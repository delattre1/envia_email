import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import os
import time

MY_ADDRESS = 'email@aqui'
PASSWORD = 'senhaaqui'

dados = pd.read_excel('aprovx.xlsx', sheet_name='Aprov')
#dados1 = dados[['Nome completo', 'Sala', 'Predio', 'E-mail']]
dados1 = dados[['Nome completo', 'E-mail']]
nomes = dados1['Nome completo'].tolist()
emails = dados1['E-mail'].tolist()

def read_template(filename):
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():
    message_template = read_template('taprov.txt')
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)
    a = 0
    # For each contact, send the email:
    #for nome, email, sala, predio in zip(nomes, emails, salas, predios):

    for nome, email in zip(nomes, emails):
        a += 1
        msg = MIMEMultipart()       # create a message
        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=nome.title())

        # Prints out the message body for our sake
        print(message)

        # setup the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="Material e Link para Aula de Contabilidade (18/05)"
        
        # add in the message body
        msg.attach(MIMEText(message, 'plain'))
        # send the message via the server set up earlier.
        s.send_message(msg)
        del msg
        time.sleep(1)
        if a == 1 or a == 80 or a == 40 or a == 20:
            time.sleep(5)
        
    # Terminate the SMTP session and close the connection
    s.quit()
    
if __name__ == '__main__':
    main()
