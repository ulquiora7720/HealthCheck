import requests
import json
from tabulate import tabulate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from requests.packages.urllib3.exceptions import InsecureRequestWarning # type: ignore

# Désactiver les avertissements de sécurité pour les requêtes HTTPS non vérifiées
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# Configuration de l'e-mail
sender_email = ""
receiver_email = ""
password = ""
smtp_server = ""
port =   # Pour starttls
subject = "Fichier Excel des serveurs critiques"


# Configuration pour se connecter à l'instance OpenManage
config = {
    "ip": "",
    "credentials": {
        "userName": "",
        "password": ""
    }
}

# URL de l'API OpenManage pour lister tous les serveurs
url = f"https://{config['ip']}/api/DeviceService/Devices?$skip=0&$top=300"

# Headers pour la requête
headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}

# Authentification
auth = (config['credentials']['userName'], config['credentials']['password'])

try:
    # Envoyer une requête GET pour récupérer les informations des périphériques
    response = requests.get(url, headers=headers, auth=auth, verify=False)

    if response.status_code == 200:
        # Convertir la réponse en JSON
        data = response.json()
        
        # Afficher la structure JSON reçue
        #print(json.dumps(data, indent=4))

        # Afficher les informations pour chaque serveur sous forme de tableau
        if 'value' in data:
            servers = data['value']
            table_data = []
            for idx, server in enumerate(servers, start=1):
                serial_number = server.get('DeviceServiceTag', 'N/A')
                model = server.get('Model', 'N/A')
                # Accéder à la liste 'DeviceManagement' s'il existe
                device_management = server.get('DeviceManagement', [])
                # Itérer sur les éléments de 'DeviceManagement' pour trouver 'NetworkAddress'
                ip = 'N/A'
                for management in device_management:
                    ip = management.get('NetworkAddress', 'N/A')
                    if ip != 'N/A':
                        break
                #ip = server.get('DeviceManagement', {}).get('NetworkAddress', 'N/A')
                status_code = server.get('Status', 5000)  # Par défaut à NOSTATUS
                status_mapping = {
                    1000: "NORMAL",
                    2000: "UNKNOWN",
                    3000: "WARNING",
                    4000: "CRITICAL",
                    5000: "NOSTATUS"
                }
                status = status_mapping.get(status_code, 'N/A')

                # Filtrer uniquement les serveurs avec le statut WARNING ou CRITICAL
                if status in ["WARNING", "CRITICAL"]:
                    table_data.append([idx, serial_number, model, ip, status])

            headers = ["#", "Numéro de série", "Modèle", "IP", "État"]
            print(tabulate(table_data, headers=headers, tablefmt="pretty"))
            server_df=pd.DataFrame(table_data,columns=[
                'idx',
                'Numéro de série',
                'Model',
                'IP',
                'Etat'
            ])
            server_df.to_excel("critical_servers.xlsx", index=False)
            # Créer le message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject
            # Ajouter le fichier Excel en pièce jointe
            filename = "critical_servers.xlsx"
            attachment = open(filename, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % filename)  
            msg.attach(part)
            # Envoyer l'e-mail
            server = smtplib.SMTP(smtp_server,port)
            server.starttls()
            server.login(sender_email, password)
            text = msg.as_string()
            server.sendmail(sender_email, receiver_email, text)
            server.quit()
        else:
            print("La clé 'value' est manquante dans la réponse JSON.")

    else:
        print(f"Erreur lors de la récupération des informations : {response.status_code}")

except requests.exceptions.RequestException as e:
    print(f"Une erreur s'est produite lors de la connexion à l'API OpenManage : {e}")
