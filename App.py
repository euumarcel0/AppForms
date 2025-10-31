import os # Para interagir com o sistema operacional (variáveis de ambiente, caminhos de arquivo)
import re # Para operações com expressões regulares (usado para extrair JSON da resposta da IA)
import json # Para manipulação de objetos JSON (usado para codificar/decodificar a resposta da IA)
import time # Para adicionar pausas no processamento (ajuda a evitar limites de taxa da API)
import tkinter as tk # Biblioteca padrão para a criação da interface gráfica (GUI)
from tkinter import filedialog, messagebox, ttk # Componentes da GUI (diálogo de arquivo, caixas de mensagem, widgets temáticos)
import threading # Para executar o processo principal em segundo plano (evita que a GUI trave)
import PyPDF2 # Biblioteca para ler e extrair texto de arquivos PDF

# Adicionado tratamento para não depender de pandas no ambiente de produção do forms
# import pandas as pd # Comentado, pois não é necessário (a manipulação de dados é feita com listas e dicionários)

from google import genai # O SDK principal do Google GenAI para interagir com o modelo Gemini
from google.genai.errors import APIError # Para capturar erros específicos da API Gemini
from google_auth_oauthlib.flow import InstalledAppFlow # Para o fluxo de autenticação OAuth 2.0 (necessário para a Google Forms API)
from googleapiclient.discovery import build # Para construir o objeto de serviço para interagir com a Google Forms API
from googleapiclient.errors import HttpError # Para capturar erros de requisições HTTP da Google Forms API (ex: erro de permissão)

# ==============================================================================
# 🔑 CONFIGURAÇÕES ESSENCIAIS
# ==============================================================================

# 1. CHAVE DA API GEMINI (Necessária para a primeira etapa: PDF -> JSON)
# ⚠️ IMPORTANTE: SUBSTITUA ESTA CHAVE PELA SUA CHAVE REAL DA API GEMINI
GEMINI_API_KEY = "chave" 

# 2. LIMITE DE TEXTO 
TEXT_LIMIT = 60000 # Limite de caracteres do texto do PDF enviado para a IA (para evitar exceder o limite do modelo)
MAX_QUESTIONS_PER_FORM = 30 # Máximo de questões que o script colocará em um único formulário do Google (limite da API ou preferência)
SCOPES = ['https://www.googleapis.com/auth/forms.body', 'https://www.googleapis.com/auth/forms.body.readonly'] 
# Escopos de permissão necessários para criar, ler e modificar o corpo de um Google Form
CREDENTIALS_FILE = 'chave.json' # Nome do arquivo de credenciais JSON do Google Cloud (para Forms API)

# Define a chave de API para a variável de ambiente (boa prática)
# O SDK do Gemini geralmente busca a chave aqui se ela não for passada explicitamente
os.environ['GEMINI_API_KEY'] = GEMINI_API_KEY


# ==============================================================================
# ⚙️ FUNÇÕES AUXILIARES
# ==============================================================================

def limpar_texto(texto):
    """
    Remove quebras de linha e espaços desnecessários de uma string.
    É essencial para limpar o texto extraído do PDF e garantir que as 
    alternativas corretas correspondam exatamente às opções.
    """
    if not isinstance(texto, str):
        texto = str(texto)
    return texto.replace('\r', ' ').replace('\n', ' ').strip()

# --- FUNÇÕES DA API GEMINI (Extração e Parsing) ---

def extract_text_from_pdf(pdf_path, progress_callback=None):
    """
    Extrai texto de um arquivo PDF usando PyPDF2.
    
    Args:
        pdf_path (str): Caminho para o arquivo PDF.
        progress_callback (function): Função para atualizar o progresso na GUI.
        
    Returns:
        str: O texto completo extraído do PDF.
    """
    text = ""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            num_pages = len(reader.pages)

            for i, page in enumerate(reader.pages):
                page_text = page.extract_text()
                if page_text:
                    # Adiciona uma quebra de linha entre as páginas para separação lógica
                    text += page_text + "\n"

                # Atualiza progresso da extração (Alocado de 0% a 10% do total)
                if progress_callback:
                    percent = int(((i + 1) / num_pages) * 10)
                    progress_callback(percent, f"1/5 - Extraindo página {i + 1} de {num_pages}...")

        return text.strip()

    except Exception as e:
        # Lança uma exceção para ser capturada na função principal
        raise Exception(f"Erro ao ler PDF: {e}")


def send_to_gemini(pdf_text, progress_callback=None):
    """
    Envia o texto do PDF para a API Gemini, solicitando uma resposta JSON estruturada.
    
    Args:
        pdf_text (str): O texto extraído do PDF.
        progress_callback (function): Função para atualizar o progresso na GUI.
        
    Returns:
        str: O texto da resposta da IA (deve conter o JSON).
    """

    if GEMINI_API_KEY == "SUA_CHAVE_AQUI" or not GEMINI_API_KEY:
        raise Exception("Chave da API Gemini ausente. Por favor, insira sua chave em GEMINI_API_KEY.")

    try:
        # Inicializa o cliente da API. O SDK usará a variável de ambiente GEMINI_API_KEY
        client = genai.Client() 
    except Exception:
        raise Exception("Erro ao inicializar o cliente Gemini. Verifique a chave de API.")

    # Limita o texto enviado ao valor de TEXT_LIMIT (60000)
    text_to_send = pdf_text[:TEXT_LIMIT]

    if progress_callback:
        progress_callback(15, "2/5 - Preparando envio para IA...")

    # O prompt detalhado é crucial para garantir que o modelo retorne um JSON estrito
    # no formato desejado para fácil parsing posterior.
    full_prompt = (
        "Analise o conteúdo extraído do simulado LPIC a seguir. "
        "Seu objetivo é extrair todas as perguntas, todas as alternativas apresentadas, "
        "e indicar a alternativa correta. O output DEVE ser um JSON estritamente válido "
        "que possa ser decodificado diretamente em uma lista (Array). "
        "Use o formato de lista de objetos JSON:\n"
        "[\n"
        "   {\n"
        "     \"numero\": 1, // número da pergunta (inteiro)\n"
        "     \"enunciado\": \"texto da pergunta\",\n"
        "     \"alternativas\": [\"opção A\", \"opção B\", \"opção C\", \"opção D\"],\n"
        "     \"correta\": \"o texto exato da alternativa correta\"\n"
        "   },\n"
        "   // ... outras perguntas\n"
        "]\n\n"
        f"CONTEÚDO DO PDF (Primeiros {len(text_to_send)} caracteres):\n\n{text_to_send}"
    )

    try:
        if progress_callback:
            progress_callback(30, "2/5 - Processando na Gemini API (aguarde)...")

        # Chama a API para geração de conteúdo
        response = client.models.generate_content(
            model="gemini-2.5-flash", # Modelo rápido e eficiente para tarefas de extração estruturada
            contents=full_prompt,
            config={"temperature": 0.1} # Temperatura baixa para respostas determinísticas (JSON estruturado)
        )

        if progress_callback:
            progress_callback(45, "2/5 - Resposta recebida da IA...")

        if not response.text:
            raise Exception("A resposta da API Gemini está vazia.")

        return response.text.strip()

    except APIError as e:
        # Tratamento específico para erros comuns da API
        if "maximum size for a single request" in str(e):
            raise Exception("Erro: O PDF é muito grande. Tente reduzir o limite de caracteres ou usar um modelo maior.")
        raise Exception(f"Erro na API Gemini: {e}")
    except Exception as e:
        raise e


def parse_gemini_response_to_list(gemini_output, progress_callback=None):
    """
    Processa a saída JSON (que pode estar envolvida em markdown) da IA 
    e retorna uma lista de dicionários de perguntas no formato final para o Forms.
    
    Args:
        gemini_output (str): A string de resposta do modelo Gemini.
        progress_callback (function): Função para atualizar o progresso na GUI.
        
    Returns:
        list: Lista de dicionários, onde cada dicionário é uma questão com 
              chaves como 'Número', 'Enunciado', 'Correta', 'A', 'B', etc.
    """

    if progress_callback:
        progress_callback(50, "3/5 - Processando resposta da IA...")

    # Tenta extrair o JSON envolto em ```json ... ``` (markdown)
    json_match = re.search(r"```json\s*(\{.*?\}|\[.*?\])\s*```", gemini_output, re.DOTALL)
    # Se falhar, tenta encontrar qualquer bloco que se pareça com JSON ({...} ou [...])
    if not json_match:
        json_match = re.search(r"(\{.*?\}|\[.*?\])", gemini_output, re.DOTALL)
    if not json_match:
        raise ValueError("Não foi possível encontrar um JSON válido na resposta da IA.")

    try:
        json_content = json_match.group(1).strip()

        # Decodifica o JSON para um objeto Python
        if json_content.startswith('{'):
            data = json.loads(json_content)
            # Trata o caso em que o modelo retorna um dicionário com uma chave 'perguntas': [...]
            if isinstance(data, dict) and any(isinstance(v, list) for v in data.values()):
                data = next((v for v in data.values() if isinstance(v, list)), [data])
            elif isinstance(data, dict):
                # Se for um único objeto de pergunta
                data = [data]
        else:
            # Caso mais comum: o JSON é uma lista diretamente
            data = json.loads(json_content)

        if not isinstance(data, list):
            raise TypeError("O JSON decodificado não é uma lista de perguntas válida.")
            
        # Converter a lista de objetos do Gemini para o formato de dicionário final
        # (com as opções A, B, C como chaves)
        processed_questions = []
        for item in data:
            alts = item.get("alternativas", [])
            q = {
                "Número": item.get("numero", ""),
                "Enunciado": item.get("enunciado", "").strip(),
                "Correta": item.get("correta", "").strip(),
            }
            # Adicionar alternativas usando letras como chaves (A, B, C, ...)
            for i in range(len(alts)):
                q[chr(65 + i)] = limpar_texto(alts[i]) # 65 é o código ASCII para 'A'
            processed_questions.append(q)

        return processed_questions

    except json.JSONDecodeError as e:
        raise ValueError(f"Erro ao decodificar JSON: {e}")


# --- FUNÇÕES DA API GOOGLE FORMS ---

def autenticar_google(progress_callback):
    """
    Autentica o usuário com o Google usando o fluxo OAuth 2.0.
    Cria as credenciais e o objeto de serviço para interagir com a Forms API.
    
    Returns:
        googleapiclient.discovery.Resource or None: O objeto de serviço da Forms API.
    """
    progress_callback(55, "4/5 - Autenticando com o Google...")
    if not os.path.exists(CREDENTIALS_FILE):
        messagebox.showerror(
            "Erro de Credenciais",
            f"Arquivo '{CREDENTIALS_FILE}' não encontrado.\nBaixe suas credenciais JSON da Google Cloud Console."
        )
        return None
    try:
        # Inicia o fluxo OAuth 2.0. Isso abrirá o navegador para o usuário fazer login.
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0) # Roda um servidor local temporário para receber o token de volta
        progress_callback(60, "4/5 - Autenticação concluída. Conectando à API...")
        # Constrói o objeto de serviço para a Forms API v1
        return build('forms', 'v1', credentials=creds)
    except Exception as e:
        messagebox.showerror("Erro de Autenticação", f"Falha ao autenticar: {e}")
        return None

def get_answer_key(question_row):
    """
    Determina o tipo de questão (RADIO para única escolha, CHECKBOX para múltipla) 
    e extrai o(s) texto(s) da(s) resposta(s) correta(s).
    
    A lógica tenta deduzir se a questão é de múltipla escolha com base no 
    enunciado (ex: "quais") ou na presença de múltiplos separadores no campo 'Correta'.
    
    Args:
        question_row (dict): Dicionário de uma única questão.
        
    Returns:
        tuple: (list de textos corretos, str com o tipo: 'RADIO' ou 'CHECKBOX').
    """
    correct_text = limpar_texto(question_row.get('Correta', ''))
    if not correct_text:
        # Se não houver resposta correta (erro na extração da IA), assume RADIO
        return None, 'RADIO'

    # Lógica para determinar CHECKBOX
    separators = [';', ' e ', ',']
    # Verifica se a resposta correta contém múltiplos separadores
    is_checkbox = any(sep in correct_text.lower() for sep in separators) and correct_text.count(' ') > 1
    
    # Se for de múltipla escolha (dedução)
    if is_checkbox or limpar_texto(question_row.get('Enunciado', '')).lower().startswith("quais"):
        question_type = 'CHECKBOX'
        correct_values = []
        # Tenta casar cada opção A, B, C... com partes do texto completo da resposta 'Correta'
        for col in [chr(65 + i) for i in range(26)]:
            option_text = limpar_texto(question_row.get(col, ''))
            # Se a opção (ex: "A") estiver contida no texto da 'Correta'
            if option_text and option_text in correct_text:
                correct_values.append(option_text)
        
        # Fallback: se não encontrou matches (o que pode acontecer se a IA 
        # formatar mal a 'Correta'), usa o texto completo como a única resposta correta
        if not correct_values:
            correct_values = [correct_text] 
        return correct_values, question_type

    # Caso padrão: escolha única (RADIO)
    question_type = 'RADIO'
    return [correct_text], question_type


def criar_forms_google(service, form_title, questions_list, progress_callback):
    """
    Cria um ou mais Forms do Google, dividindo as questões em lotes de 
    MAX_QUESTIONS_PER_FORM. Para cada Forms, ativa o modo Quiz e adiciona as questões.
    
    Args:
        service (Resource): O objeto de serviço da Google Forms API.
        form_title (str): Título base do formulário.
        questions_list (list): Lista de dicionários de questões extraídas.
        progress_callback (function): Função para atualizar o progresso na GUI.
        
    Returns:
        tuple: (list de links dos Forms criados, int total de questões).
    """

    # Calcula quantos Forms serão necessários
    num_forms = (len(questions_list) + MAX_QUESTIONS_PER_FORM - 1) // MAX_QUESTIONS_PER_FORM
    all_form_links = []
    
    # Define a faixa de progresso para esta etapa (65% a 100%)
    PROGRESS_RANGE_START = 65
    PROGRESS_RANGE_END = 100
    TOTAL_PROGRESS_POINTS = PROGRESS_RANGE_END - PROGRESS_RANGE_START
    
    # Itera sobre os lotes de questões para criar múltiplos Forms
    for i in range(num_forms):
        start = i * MAX_QUESTIONS_PER_FORM
        end = min(len(questions_list), start + MAX_QUESTIONS_PER_FORM)
        # O FATIAMENTO É FEITO AQUI: Seleciona o lote de questões
        part_questions = questions_list[start:end] 
        
        # Cálculo de progresso para este formulário
        progress_per_form = TOTAL_PROGRESS_POINTS / num_forms
        current_form_start_progress = PROGRESS_RANGE_START + (i * progress_per_form)
        
        # Título personalizado para cada parte
        title = f"{form_title} - Parte {i + 1} ({len(part_questions)} Q)"
        
        # --- 1. Criar Forms e ativar Quiz ---
        try:
            progress_callback(int(current_form_start_progress), f"5/5 - Criando Forms {i + 1}/{num_forms}...")
            # Cria o formulário com o título
            form = service.forms().create(body={'info': {'title': limpar_texto(title)}}).execute()
            form_id = form['formId']
        except HttpError as e:
            messagebox.showerror("Erro de Criação", f"Não foi possível criar o Forms: {e}")
            continue

        # Requisita a atualização para ativar o modo Quiz no Forms
        service.forms().batchUpdate(
            formId=form_id,
            body={'requests': [{'updateSettings': {'settings': {'quizSettings': {'isQuiz': True}}, 'updateMask': 'quizSettings.isQuiz'}}]}
        ).execute()

        # --- 2. Preparar Requisições Batch ---
        requests = []
        index = 0
        
        # Itera sobre as questões do lote para montar as requisições de criação
        for question_row in part_questions: # Usando a lista FATIADA
            title_text = limpar_texto(f"Q{str(question_row.get('Número', '')).strip()}: {question_row.get('Enunciado', '')}")
            correct_values, question_type = get_answer_key(question_row)

            options = []
            option_cols = [col for col in question_row.keys() if len(col) == 1 and 'A' <= col <= 'Z']
            option_set = set() # Usado para evitar opções duplicadas

            for col in option_cols:
                option_text = question_row.get(col, '')
                if not option_text or not str(option_text).strip():
                    continue
                option_text = limpar_texto(option_text)
                if option_text not in option_set:
                    options.append({'value': option_text})
                    option_set.add(option_text)

            if not options: continue

            # Filtra as opções que correspondem ao texto correto
            answer_key_texts = [opt['value'] for opt in options if correct_values and opt['value'] in correct_values]

            # Objeto de pontuação (Grading)
            grading = {
                'pointValue': 1,
                'correctAnswers': {'answers': [{'value': v} for v in answer_key_texts]}
            } if answer_key_texts else None # Se não houver resposta correta, não adiciona 'grading'

            # Corpo da questão (Question Body)
            question_body = {
                'required': True,
                'choiceQuestion': {
                    'type': question_type, # 'RADIO' ou 'CHECKBOX'
                    'options': options,
                    'shuffle': True # Misturar a ordem das opções
                }
            }
            if grading: question_body['grading'] = grading

            # Requisição de criação de item
            requests.append({
                'createItem': {
                    'item': {'title': title_text, 'questionItem': {'question': question_body}},
                    'location': {'index': index}
                }
            })
            index += 1
        
        # --- 3. Enviar Requisições em Lotes e Atualizar Progresso ---
        created_count = 0
        total_requests = len(requests)
        # Divide as requisições em lotes de 10 para otimizar a chamada à API (Batch Update)
        for j in range(0, total_requests, 10):
            batch = requests[j:j + 10]
            try:
                service.forms().batchUpdate(formId=form_id, body={'requests': batch}).execute()
                created_count += len(batch)
                
                # Cálculo de progresso dentro do Form atual
                progress_in_form = (created_count / total_requests) * progress_per_form
                current_overall_progress = current_form_start_progress + progress_in_form
                
                progress_callback(int(current_overall_progress), f"5/5 - Criando Forms {i + 1}/{num_forms}: {created_count}/{total_requests} questões...")
                time.sleep(0.3) # Pequena pausa para evitar sobrecarga
            except HttpError as e:
                print(f"⚠️ Erro ao adicionar lote {j//10+1} ao Forms {i+1}: {e}")
                continue

        link = f"https://docs.google.com/forms/d/{form_id}/edit"
        print(f"✅ Formulário '{title}' criado ({created_count} questões). Link: {link}")
        all_form_links.append(link)

    # Última atualização de progresso
    progress_callback(PROGRESS_RANGE_END, "5/5 - Criação de Forms concluída.") 
    return all_form_links, len(questions_list)


# --- LÓGICA PRINCIPAL E UI ---

class PipelineApp:
    """
    Classe principal que gerencia a Interface Gráfica (GUI) e a execução do pipeline.
    """
    def __init__(self, master):
        self.master = master
        master.title("LPIC PDF → IA (Gemini) → Google Forms")
        master.geometry("450x250") # Tamanho fixo da janela
        master.resizable(False, False) # Impede redimensionamento

        # Configuração de estilo para a barra de progresso
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("blue.Horizontal.TProgressbar", foreground='#3B82F6', background='#3B82F6')

        # Título e instruções da aplicação
        tk.Label(
            master,
            text="Pipeline: Selecione um PDF → Extração IA (Gemini) → Google Forms.",
            pady=15,
            padx=20,
            wraplength=400,
            justify="center",
            font=('Arial', 10, 'bold')
        ).pack()

        # Botão principal para iniciar o processo
        self.btn_start = tk.Button(
            master,
            text="📂 Selecionar PDF e Criar Forms",
            command=self.run_process_in_thread, # Chama a função que inicia o processo em uma nova thread
            padx=20,
            pady=10,
            bg="#3B82F6",
            fg="white"
        )
        self.btn_start.pack(pady=10)

        # Barra de progresso (determinate = mostra o progresso de 0 a 100)
        self.progress_bar = ttk.Progressbar(
            master,
            orient='horizontal',
            length=400,
            mode='determinate',
            style="blue.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=10)

        # Rótulo de status (mostra a etapa atual)
        self.status_label = tk.Label(master, text="Aguardando início...", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X)

    def update_progress(self, value, text):
        """
        Atualiza a barra de progresso e o rótulo de status na thread principal da GUI.
        'self.master.after(0, ...)' garante que a atualização ocorra de forma segura.
        """
        self.master.after(0, lambda: [
            self.progress_bar.config(value=value),
            self.status_label.config(text=text),
            self.master.update_idletasks() # Força a atualização da interface
        ])

    def run_creation_logic(self):
        """
        Função que contém a lógica completa do pipeline, executada em uma thread separada.
        Gerencia o fluxo de trabalho e o tratamento de erros.
        """
        
        # Checagem inicial da chave de API
        if GEMINI_API_KEY == "SUA_CHAVE_AQUI" or not GEMINI_API_KEY:
             messagebox.showwarning(
                 "Chave da API ausente",
                 "⚠️ Por favor, edite o código e insira sua chave da API Gemini na variável GEMINI_API_KEY."
             )
             self.btn_start.config(state=tk.NORMAL)
             return
             
        self.update_progress(0, "Iniciando o processo...")
        self.btn_start.config(state=tk.DISABLED) # Desabilita o botão para evitar cliques múltiplos

        # Diálogo para selecionar o arquivo PDF
        pdf_path = filedialog.askopenfilename(
            title="Selecione o PDF do simulado LPIC",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not pdf_path:
            self.btn_start.config(state=tk.NORMAL)
            self.update_progress(0, "Processo cancelado.")
            return

        try:
            # 1. Extrair texto do PDF (0% a 10%)
            raw_text = extract_text_from_pdf(pdf_path, self.update_progress)
            
            # 2. Enviar para Gemini (10% a 50%)
            gemini_response = send_to_gemini(raw_text, self.update_progress)

            # 3. Processar resposta da Gemini (50% a 55%)
            questions_list = parse_gemini_response_to_list(gemini_response, self.update_progress)
            
            # 4. Autenticar Google Forms (55% a 65%)
            service = autenticar_google(self.update_progress)
            if not service:
                return # Retorna se a autenticação falhar

            # 5. Criar Google Forms (65% a 100%)
            form_title_base = os.path.basename(pdf_path).replace('.pdf', '')
            form_links, num_questions = criar_forms_google(
                service, 
                form_title_base, 
                questions_list, 
                self.update_progress 
            )

            # 6. Exibir sucesso
            self.update_progress(100, "Concluído com sucesso!")
            messagebox.showinfo(
                "Sucesso",
                f"✅ Extraídas e Criadas {num_questions} questões em {len(form_links)} Forms(s).\n\n"
                f"Links dos Forms:\n{'\n'.join(form_links)}"
            )

        except Exception as e:
            # Captura e exibe qualquer erro ocorrido em qualquer etapa
            self.update_progress(0, "Erro: " + str(e))
            messagebox.showerror("Erro", f"Falha no processamento:\n{str(e)}")

        finally:
            # Bloco executado sempre, reabilita o botão e zera a barra de progresso
            self.btn_start.config(state=tk.NORMAL)
            self.progress_bar.config(value=0)

    def run_process_in_thread(self):
        """
        Inicia a função run_creation_logic em uma thread separada.
        Isso é crucial para que a Interface Gráfica (GUI) permaneça responsiva 
        enquanto as chamadas de API de longa duração (Gemini e Forms) estão em execução.
        """
        threading.Thread(target=self.run_creation_logic).start()


if __name__ == '__main__':
    # Bloco de execução principal da aplicação
    root = tk.Tk()
    app = PipelineApp(root)
    root.mainloop() # Inicia o loop principal da GUI