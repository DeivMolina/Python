import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Leer el archivo Excel
workbook = openpyxl.load_workbook('datos.xlsx')  # Reemplaza 'datos.xlsx' con el nombre de tu archivo Excel
sheet = workbook.active

# Configuración del servidor SMTP
smtp_server = 'mail.panduitlatam.com'
smtp_port = 465
smtp_user = 'info@panduitlatam.com'
smtp_password = 'rIu6Q.Ts&B~o'

# Dirección de correo para copia oculta (CCO)
cco_address = 'wb.rkcreativo@gmail.com'

# Conexión al servidor SMTP
server = smtplib.SMTP_SSL(smtp_server, smtp_port)
server.login(smtp_user, smtp_password)

# Contador para controlar el número de correos enviados
num_correos_enviados = 0

# Iterar sobre las filas de la hoja de cálculo
for i, row in enumerate(sheet.iter_rows(values_only=True)):
    nombre_archivo = row[0]  # Nombre del archivo HTML
    destinatario = row[1]  # Dirección de correo electrónico del destinatario
    
    # Configurar el mensaje
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg['Subject'] = '🏅 Reconocimiento de Participación Panduit Week'
    
    # Agregar CCO
    msg['Bcc'] = cco_address
    
    # Leer el contenido HTML del archivo
    file_path = f'output_html/{nombre_archivo}.html'
    print("Ruta del archivo:", file_path)
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    # Adjuntar el contenido HTML al mensaje
    html_part = MIMEText(html_content, 'html')
    msg.attach(html_part)
    
    # Enviar el correo electrónico
    server.send_message(msg)

    # Incrementar el contador de correos enviados
    num_correos_enviados += 1

    # Cerrar y abrir la conexión SMTP después de enviar cada 10 correos electrónicos
    if num_correos_enviados % 10 == 0:
        server.quit()
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)  # Aumenta el valor de timeout según sea necesario
        server.login(smtp_user, smtp_password)

# Cerrar conexión SMTP al finalizar
server.quit()
