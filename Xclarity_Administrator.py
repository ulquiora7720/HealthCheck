import requests
import json
from tabulate import tabulate
import pandas as pd
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

# Configuration pour se connecter à l'instance XClarity
config = {
    "ip": "",
    "credentials": {
        "userName": "",
        "password": ""
    }
}

# URL de l'API XClarity pour lister tous les serveurs
url = f"https://{config['ip']}/nodes"

# Headers pour la requête
headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}

# Authentification
auth = (config['credentials']['userName'], config['credentials']['password'])

# Mapping des statuts
status_mapping = {
    "Normal": "NORMAL",
    "Major-Failure": "CRITICAL",
    "Warning": "WARNING"
}

# Création d'une session pour réutiliser les connexions
with requests.Session() as session:
    session.headers.update(headers)
    session.auth = auth

    try:
        # Envoyer une requête GET pour récupérer les informations des périphériques
        response = session.get(url, verify=False)
        response.raise_for_status()  # Raise an HTTPError for bad responses

        # Convertir la réponse en JSON
        data = response.json()

        # Afficher la structure JSON reçue
        #print(json.dumps(data, indent=4))

        # Afficher les informations pour chaque serveur sous forme de tableau
        if 'nodeList' in data:
            servers = data['nodeList']
            table_data = []
            for idx, server in enumerate(servers, start=1):
                serial_number = server.get('serialNumber', 'N/A')
                model = server.get('machineType', 'N/A')
                ip = server.get('mgmtProcIPaddress', 'N/A')
                overall_health = server.get('overallHealthState', 'N/A')
                status = status_mapping.get(overall_health, 'N/A')

                # Filter only servers with WARNING or CRITICAL status
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
            # Affichage avec pandas pour plus de flexibilité
            df = pd.DataFrame(table_data, columns=["#", "Serial Number", "Model", "IP", "Status"])
        else:
            print("The 'nodeList' key is missing in the JSON response.")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred while connecting to the XClarity API: {e}")
    except json.JSONDecodeError:
        print("Failed to decode JSON response.")
    except KeyError as e:
        print(f"Missing expected key in the JSON response: {e}")
