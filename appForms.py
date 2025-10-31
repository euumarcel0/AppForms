import pandas as pd
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configurações gerais
SCOPES = ['https://www.googleapis.com/auth/forms.body', 'https://www.googleapis.com/auth/forms.body.readonly']
CREDENTIALS_FILE = 'chave.json'
MAX_QUESTIONS_PER_FORM = 30


def limpar_texto(texto):
    """Remove quebras de linha e espaços desnecessários."""
    if not isinstance(texto, str):
        texto = str(texto)
    return texto.replace('\r', ' ').replace('\n', ' ').strip()


class FormsCreatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Criador de Google Forms Automatizado")
        master.geometry("450x250")
        master.resizable(False, False)

        self.service = None
        self.excel_file = None

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("blue.Horizontal.TProgressbar", foreground='#3B82F6', background='#3B82F6')

        tk.Label(
            master,
            text="Selecione o arquivo Excel extraído para criar o(s) Google Forms.",
            pady=15,
            padx=20,
            wraplength=400,
            justify="center",
            font=('Arial', 10, 'bold')
        ).pack()

        self.btn_start = tk.Button(
            master,
            text="📂 Selecionar Excel e Criar Forms",
            command=self.run_process_in_thread,
            padx=20,
            pady=10,
            bg="#4CAF50",
            fg="white"
        )
        self.btn_start.pack(pady=10)

        self.progress_bar = ttk.Progressbar(
            master,
            orient='horizontal',
            length=400,
            mode='determinate',
            style="blue.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=10)

        self.status_label = tk.Label(master, text="Aguardando início...", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X)

    def update_progress(self, value, text):
        self.master.after(0, lambda: [
            self.progress_bar.config(value=value),
            self.status_label.config(text=text),
            self.master.update_idletasks()
        ])

    def autenticar_google(self):
        self.update_progress(10, "1/5 - Autenticando com o Google...")
        if not os.path.exists(CREDENTIALS_FILE):
            messagebox.showerror(
                "Erro de Credenciais",
                f"Arquivo '{CREDENTIALS_FILE}' não encontrado.\nBaixe suas credenciais JSON da Google Cloud Console."
            )
            return None
        try:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            self.update_progress(30, "2/5 - Autenticação concluída. Conectando à API...")
            return build('forms', 'v1', credentials=creds)
        except Exception as e:
            messagebox.showerror("Erro de Autenticação", f"Falha ao autenticar: {e}")
            return None

    def get_answer_key(self, question_row):
        correct_text = limpar_texto(question_row.get('Correta', ''))
        if not correct_text:
            return None, 'RADIO'

        separators = [';', ' e ', ',']
        is_checkbox = any(sep in correct_text.lower() for sep in separators) and correct_text.count(' ') > 1

        if is_checkbox or limpar_texto(question_row.get('Enunciado', '')).lower().startswith("quais"):
            question_type = 'CHECKBOX'
            correct_values = []
            for col in [chr(65 + i) for i in range(26)]:
                option_text = limpar_texto(question_row.get(col, ''))
                if pd.notna(option_text) and option_text and option_text in correct_text:
                    correct_values.append(option_text)
            if not correct_values:
                correct_values = [correct_text]
            return correct_values, question_type

        question_type = 'RADIO'
        return [correct_text], question_type

    def criar_forms_google(self, service, form_title, questions_df, form_total_start_progress, form_total_end_progress):
        try:
            form = service.forms().create(body={'info': {'title': limpar_texto(form_title)}}).execute()
            form_id = form['formId']
        except HttpError as e:
            messagebox.showerror("Erro de Criação", f"Não foi possível criar o Forms: {e}")
            return None, 0

        # Ativar modo quiz
        service.forms().batchUpdate(
            formId=form_id,
            body={'requests': [{
                'updateSettings': {
                    'settings': {'quizSettings': {'isQuiz': True}},
                    'updateMask': 'quizSettings.isQuiz'
                }
            }]}
        ).execute()

        # Dentro da função criar_forms_google
        requests = []
        index = 0  # índice crescente

        for _, question_row in questions_df.iterrows():
            title_text = limpar_texto(f"Q{str(question_row.get('Número', '')).strip()}: {question_row.get('Enunciado', '')}")
            correct_values, question_type = self.get_answer_key(question_row)

            options = []
            option_cols = [col for col in question_row.index if len(col) == 1 and 'A' <= col <= 'Z']
            option_set = set()

            for col in option_cols:
                option_text = question_row.get(col, '')
                if pd.isna(option_text) or not str(option_text).strip():
                    continue  # pula valores vazios ou NaN
                option_text = limpar_texto(option_text)
                if option_text not in option_set:
                    options.append({'value': option_text})
                    option_set.add(option_text)

            if not options:
                continue  # pula questões sem opções válidas

            answer_key_texts = [opt['value'] for opt in options if correct_values and opt['value'] in correct_values]

            grading = {
                'pointValue': 1,
                'correctAnswers': {'answers': [{'value': v} for v in answer_key_texts]}
            } if answer_key_texts else None

            question_body = {
                'required': True,
                'choiceQuestion': {
                    'type': question_type,
                    'options': options,
                    'shuffle': True
                }
            }
            if grading:
                question_body['grading'] = grading

            requests.append({
                'createItem': {
                    'item': {
                        'title': title_text,
                        'questionItem': {'question': question_body}
                    },
                    'location': {'index': index}  # índice crescente
                }
            })
            index += 1  # incrementa para próxima pergunta

        created_count = 0
        total_requests = len(requests)

        # Envia em blocos maiores (10 por vez) e continua mesmo se alguma falhar
        for i in range(0, total_requests, 10):
            batch = requests[i:i + 10]
            try:
                service.forms().batchUpdate(formId=form_id, body={'requests': batch}).execute()
                created_count += len(batch)
                progress = form_total_start_progress + (created_count / total_requests) * (form_total_end_progress - form_total_start_progress)
                self.update_progress(progress, f"Adicionando questões: {created_count}/{total_requests}...")
                time.sleep(0.3)
            except HttpError as e:
                print(f"⚠️ Erro ao adicionar lote {i//10+1}: {e}")
                continue  # Não para o processo — apenas pula o lote problemático

        return form_id, created_count

    def run_creation_logic(self):
        self.btn_start.config(state=tk.DISABLED)
        self.progress_bar.config(value=0)

        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel de Questões",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if not file_path:
            self.btn_start.config(state=tk.NORMAL)
            return

        self.service = self.autenticar_google()
        if not self.service:
            self.btn_start.config(state=tk.NORMAL)
            return

        df = pd.read_excel(file_path)
        df = df[df['Enunciado'].notna()]
        total = len(df)
        if total == 0:
            messagebox.showinfo("Aviso", "Nenhuma questão válida encontrada.")
            self.btn_start.config(state=tk.NORMAL)
            return

        num_forms = (total + MAX_QUESTIONS_PER_FORM - 1) // MAX_QUESTIONS_PER_FORM
        form_links = []

        for i in range(num_forms):
            start = i * MAX_QUESTIONS_PER_FORM
            end = min(total, start + MAX_QUESTIONS_PER_FORM)
            part_df = df.iloc[start:end]
            title = f"{os.path.basename(file_path).replace('.xlsx','')} - Parte {i + 1} ({len(part_df)} Q)"
            form_id, created = self.criar_forms_google(self.service, title, part_df, 40, 90)
            if form_id:
                link = f"https://docs.google.com/forms/d/{form_id}/edit"
                print(f"✅ Formulário '{title}' criado ({created} questões). Link: {link}")
                form_links.append(link)

        self.update_progress(100, "Processo concluído com sucesso ✅")
        if form_links:
            messagebox.showinfo("Sucesso", "\n".join(form_links))
        self.btn_start.config(state=tk.NORMAL)

    def run_process_in_thread(self):
        threading.Thread(target=self.run_creation_logic).start()


if __name__ == '__main__':
    root = tk.Tk()
    app = FormsCreatorApp(root)
    root.mainloop()
