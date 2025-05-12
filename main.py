import flet as ft
import imaplib
import email
from email.header import decode_header
import pandas as pd
import re
from threading import Thread

def main(page: ft.Page):
    page.title = "Extrair Palavras-chave de Emails"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.window_width = 600
    page.window_height = 600

    # Elementos da UI
    email_field = ft.TextField(label="Email", width=500)
    password_field = ft.TextField(label="Senha", password=True, width=500)
    keywords_field = ft.TextField(
        label="Palavras-chave (separadas por vírgula)",
        width=500
    )
    status_text = ft.Text()
    progress_ring = ft.ProgressRing(visible=False)

    def update_status(msg, error=False):
        status_text.value = msg
        status_text.color = ft.colors.RED if error else ft.colors.BLACK
        progress_ring.visible = False
        page.update()

    def process_emails(e):
        user_email = email_field.value
        password = password_field.value
        keywords = [k.strip().lower() for k in keywords_field.value.split(',')]

        if not all([user_email, password, keywords]):
            update_status("Preencha todos os campos!", error=True)
            return

        def thread_proc():
            try:
                # Configuração IMAP para Outlook
                mail = imaplib.IMAP4_SSL('outlook.office365.com')
                mail.login(user_email, password)
                mail.select('inbox')

                _, data = mail.search(None, 'ALL')
                email_ids = data[0].split()
                emails_data = []

                for i, eid in enumerate(email_ids[:20]):  # Limitar a 20 emails
                    _, msg_data = mail.fetch(eid, '(RFC822)')
                    msg = email.message_from_bytes(msg_data[0][1])

                    # Decodificar cabeçalhos
                    subject = decode_header(msg['Subject'])[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode()

                    from_ = msg.get('From')
                    date_ = msg.get('Date')

                    # Extrair texto do corpo
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == 'text/plain':
                                body = part.get_payload(decode=True).decode(errors='ignore')
                                break
                    else:
                        body = msg.get_payload(decode=True).decode(errors='ignore')

                    # Processar texto
                    clean_body = re.sub(r'\s+', ' ', body).lower()
                    keywords_count = {keyword: clean_body.count(keyword) for keyword in keywords}

                    emails_data.append({
                        'Assunto': subject,
                        'Remetente': from_,
                        'Data': date_,
                        **keywords_count
                    })

                    # Atualizar progresso
                    page.run_task(
                        lambda: status_text.update(
                            value=f"Processando {i+1}/{len(email_ids[:20])} emails..."
                        )
                    )

                # Criar DataFrame e salvar Excel
                df = pd.DataFrame(emails_data)
                df.to_excel('emails_processados.xlsx', index=False)
                page.run_task(lambda: update_status("Planilha gerada com sucesso!"))

            except Exception as e:
                page.run_task(lambda: update_status(f"Erro: {str(e)}", error=True))
            finally:
                mail.logout()

        # Mostrar indicador de progresso
        progress_ring.visible = True
        page.update()

        # Iniciar thread
        Thread(target=thread_proc).start()

    page.add(
        ft.Column([
            ft.Image(src="https://img.icons8.com/clouds/200/email.png"),
            email_field,
            password_field,
            keywords_field,
            ft.ElevatedButton(
                "Processar Emails",
                icon=ft.icons.EMAIL,
                on_click=process_emails
            ),
            progress_ring,
            status_text
        ], alignment=ft.MainAxisAlignment.CENTER)
    )

ft.app(target=main)