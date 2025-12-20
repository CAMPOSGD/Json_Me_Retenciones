import imaplib
import email
import os
from email.header import decode_header

# --- CONFIGURACIÓN ---
# Ejemplo para GMAIL. Si es Outlook usa: 'outlook.office365.com'
IMAP_SERVER = "imap.gmail.com" 
EMAIL_USER = "tu_correo@gmail.com"
EMAIL_PASS = "tu_contraseña_de_aplicacion" # NO la normal

# Carpeta donde caerán los archivos
CARPETA_DESTINO = r"C:\Users\gcampos\OneDrive\Development\Json-Me\Json"

def descargar_adjuntos():
    # 1. Conexión al servidor
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox") # O la carpeta donde llegan los correos

    # 2. Buscar correos (Filtramos por asunto para no leer todo)
    # Ejemplo: Buscamos correos que digan "Retencion" en el asunto
    status, messages = mail.search(None, '(SUBJECT "Retencion")')
    
    email_ids = messages[0].split()

    print(f"Se encontraron {len(email_ids)} correos.")

    for email_id in email_ids:
        # Traer el correo
        res, msg = mail.fetch(email_id, "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                
                # Decodificar el asunto para mostrarlo en el log
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                
                print(f"Procesando: {subject}")

                # 3. Iterar sobre las partes del correo (adjuntos)
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue

                    file_name = part.get_filename()

                    if file_name:
                        # Decodificar nombre de archivo si viene raro
                        fn, encoding = decode_header(file_name)[0]
                        if isinstance(fn, bytes):
                            file_name = fn.decode(encoding if encoding else "utf-8")

                        # 4. FILTRO: Solo descargamos JSON y PDF
                        if file_name.lower().endswith((".json", ".pdf")):
                            filepath = os.path.join(CARPETA_DESTINO, file_name)
                            
                            # Verificamos si ya existe para no sobrescribir
                            if not os.path.isfile(filepath):
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"   -> Descargado: {file_name}")
                            else:
                                print(f"   -> Ya existe: {file_name}")

    mail.close()
    mail.logout()

# Ejecutar
if __name__ == "__main__":
    descargar_adjuntos()