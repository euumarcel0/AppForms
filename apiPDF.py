import os
import re
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import PyPDF2
import pandas as pd
from google import genai
from google.genai.errors import APIError

# 🔑 SUBSTITUA PELA SUA CHAVE DA API GEMINI
GEMINI_API_KEY = "chave"  # <-- ALTERE ISSO!

# Define a chave de API para a variável de ambiente (boa prática)
os.environ['GEMINI_API_KEY'] = GEMINI_API_KEY

# Aumentamos o limite de caracteres para capturar todas as 60 questões
TEXT_LIMIT = 60000


def extract_text_from_pdf(pdf_path, progress_callback=None):
    """Extrai texto de um arquivo PDF, simulando progresso por página."""
    text = ""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            num_pages = len(reader.pages)

            for i, page in enumerate(reader.pages):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

                # Atualiza progresso da extração
                if progress_callback:
                    percent = int(((i + 1) / num_pages) * 50)  # 50% para extração
                    progress_callback(percent, f"Extraindo página {i + 1} de {num_pages}...")

        return text.strip()

    except Exception as e:
        raise Exception(f"Erro ao ler PDF: {e}")


def send_to_gemini(pdf_text, progress_callback=None):
    """Envia o texto do PDF para a API Gemini para extração estruturada."""

    try:
        client = genai.Client()
    except Exception:
        raise Exception("Erro ao inicializar o cliente Gemini. Verifique a chave de API.")

    text_to_send = pdf_text[:TEXT_LIMIT]

    if progress_callback:
        progress_callback(50, "Preparando envio para IA...")

    full_prompt = (
        "Analise o conteúdo extraído do simulado LPIC a seguir. "
        "Seu objetivo é extrair todas as perguntas, todas as alternativas apresentadas, "
        "e indicar a alternativa correta. O output DEVE ser um JSON estritamente válido. "
        "Use o formato de lista de objetos JSON:\n"
        "[\n"
        "  {\n"
        "    \"numero\": 1, // número da pergunta (inteiro)\n"
        "    \"enunciado\": \"texto da pergunta\",\n"
        "    \"alternativas\": [\"opção A\", \"opção B\", \"opção C\", \"opção D\"],\n"
        "    \"correta\": \"o texto exato da alternativa correta\"\n"
        "  },\n"
        "  // ... outras perguntas\n"
        "]\n\n"
        f"CONTEÚDO DO PDF (Primeiros {len(text_to_send)} caracteres):\n\n{text_to_send}"
    )

    try:
        if progress_callback:
            progress_callback(75, "Processando na Gemini API (aguarde)...")

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=full_prompt,
            config={"temperature": 0.1}
        )

        if progress_callback:
            progress_callback(95, "Resposta recebida...")

        if not response.text:
            raise Exception("A resposta da API Gemini está vazia.")

        return response.text.strip()

    except APIError as e:
        if "maximum size for a single request" in str(e):
            raise Exception("Erro: O PDF é muito grande. Tente reduzir o limite de caracteres ou usar um modelo maior.")
        raise Exception(f"Erro na API Gemini: {e}")
    except Exception as e:
        raise e


def parse_gemini_response_to_excel(gemini_output, output_excel, progress_callback=None):
    """Processa a saída JSON da IA e salva em um arquivo Excel."""

    if progress_callback:
        progress_callback(97, "Convertendo para Excel...")

    json_match = re.search(r"```json\s*(\{.*?\}|\[.*?\])\s*```", gemini_output, re.DOTALL)
    if not json_match:
        json_match = re.search(r"(\{.*?\}|\[.*?\])", gemini_output, re.DOTALL)
    if not json_match:
        raise ValueError("Não foi possível encontrar um JSON válido na resposta da IA.")

    try:
        data = json.loads(json_match.group(1))
        if isinstance(data, dict) and "perguntas" in data:
            data = data["perguntas"]
        elif isinstance(data, dict):
            data = [data]
        elif not isinstance(data, list):
            raise TypeError("O JSON decodificado não é uma lista de perguntas.")
    except json.JSONDecodeError as e:
        raise ValueError(f"Erro ao decodificar JSON: {e}")

    rows = []
    for item in data:
        alts = item.get("alternativas", [])
        row = {
            "Número": item.get("numero", ""),
            "Enunciado": item.get("enunciado", "").strip(),
        }

        for i in range(len(alts)):
            row[chr(65 + i)] = alts[i]

        row["Correta"] = item.get("correta", "").strip()
        rows.append(row)

    df = pd.DataFrame(rows)
    df.to_excel(output_excel, index=False)

    if progress_callback:
        progress_callback(100, "Concluído!")

    return len(rows)


def process_with_gemini(root, btn, progress_bar, status_label):
    """Função principal com a lógica de extração e API, atualizando a UI."""

    if GEMINI_API_KEY == "SUA_CHAVE_AQUI" or not GEMINI_API_KEY:
        messagebox.showwarning(
            "Chave da API ausente",
            "⚠️ Por favor, edite o código e insira sua chave da API Gemini na variável GEMINI_API_KEY."
        )
        return

    btn.config(state=tk.DISABLED)
    progress_bar.stop()
    progress_bar['value'] = 0
    status_label.config(text="Aguardando seleção do arquivo...")

    pdf_path = filedialog.askopenfilename(
        title="Selecione o PDF do simulado LPIC",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not pdf_path:
        btn.config(state=tk.NORMAL)
        status_label.config(text="Processo cancelado.")
        return

    def update_progress(value, text):
        progress_bar['value'] = value
        status_label.config(text=text)
        root.update_idletasks()  # Força atualização da interface

    try:
        # 1. Extrair texto do PDF
        update_progress(0, "Iniciando extração do PDF...")
        raw_text = extract_text_from_pdf(pdf_path, update_progress)
        if not raw_text:
            messagebox.showerror("Erro", "Não foi possível extrair texto do PDF.")
            return

        # 2. Enviar para Gemini
        gemini_response = send_to_gemini(raw_text, update_progress)

        # 3. Salvar resposta bruta em .txt
        update_progress(96, "Salvando resposta da IA...")
        txt_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt")],
            title="Salvar resposta bruta da IA (JSON) como..."
        )
        if not txt_path:
            return

        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(gemini_response)

        # 4. Gerar Excel
        excel_path = txt_path.replace(".txt", "_questoes.xlsx")
        num_questions = parse_gemini_response_to_excel(gemini_response, excel_path, update_progress)

        # 5. Exibir sucesso
        messagebox.showinfo(
            "Sucesso",
            f"✅ Extraídas {num_questions} questões.\n"
            f"Resposta bruta salva em:\n{txt_path}\n"
            f"Excel gerado em:\n{excel_path}"
        )

    except Exception as e:
        update_progress(0, "Erro: " + str(e))
        messagebox.showerror("Erro", f"Falha no processamento:\n{str(e)}")

    finally:
        btn.config(state=tk.NORMAL)
        progress_bar['value'] = 0


# --- Interface Gráfica ---
root = tk.Tk()
root.title("LPIC PDF → IA (Gemini) → Excel")
root.geometry("450x230")
root.resizable(False, False)

# Estilo da barra de progresso
style = ttk.Style()
style.theme_use('clam')
style.configure("green.Horizontal.TProgressbar", foreground='#4CAF50', background='#4CAF50')

# Label de instrução
label = tk.Label(
    root,
    text="Selecione um PDF de simulado LPIC para extrair perguntas com IA (Gemini)",
    pady=15,
    padx=20,
    wraplength=400,
    justify="center"
)
label.pack()

# Botão principal
btn = tk.Button(
    root,
    text="📁 Selecionar PDF e Processar com IA (Gemini)",
    command=lambda: process_with_gemini(root, btn, progress_bar, status_label),
    padx=20,
    pady=10,
    bg="#3B82F6",
    fg="white"
)
btn.pack(pady=5)

# Barra de progresso
progress_bar = ttk.Progressbar(
    root,
    orient='horizontal',
    length=400,
    mode='determinate',
    style="green.Horizontal.TProgressbar"
)
progress_bar.pack(pady=5)

# Rótulo de status
status_label = tk.Label(
    root,
    text="Aguardando seleção do arquivo...",
    bd=1,
    relief=tk.SUNKEN,
    anchor=tk.W
)
status_label.pack(fill=tk.X)

root.mainloop()
