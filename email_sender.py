from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os


# Zugriffsrechte für Gmail API
BEREICHE = ['https://www.googleapis.com/auth/gmail.send']


def authentifizieren():
    """Authentifiziert den Benutzer und gibt die Anmeldeinformationen zurück."""
    anmeldeinformationen = None
    # Überprüfen, ob ein Token gespeichert ist
    if os.path.exists('token.json'):
        anmeldeinformationen = Credentials.from_authorized_user_file('token.json', BEREICHE)

    # Wenn das Token ungültig ist, führen wir die Authentifizierung durch
    if not anmeldeinformationen or not anmeldeinformationen.valid:
        if anmeldeinformationen and anmeldeinformationen.expired and anmeldeinformationen.refresh_token:
            anmeldeinformationen.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', BEREICHE)
            anmeldeinformationen = flow.run_local_server(port=0)
        # Token für zukünftige Nutzung speichern
        with open('token.json', 'w') as token_datei:
            token_datei.write(anmeldeinformationen.to_json())

    return anmeldeinformationen


def erstelle_email_mit_anhang(empfaenger, betreff, nachricht, anhang_pfad):
    """Erstellt eine MIME-Email mit Anhang."""
    # Email erstellen
    email = MIMEMultipart()
    email['to'] = empfaenger
    email['subject'] = betreff
    email.attach(MIMEText(nachricht, 'plain'))

    # Anhang hinzufügen
    if os.path.exists(anhang_pfad):
        with open(anhang_pfad, 'rb') as anhang:
            teil = MIMEBase('application', 'octet-stream')
            teil.set_payload(anhang.read())
            encoders.encode_base64(teil)
            teil.add_header(
                'Content-Disposition',
                f'attachment; filename="{os.path.basename(anhang_pfad)}"'
            )
            email.attach(teil)
    else:
        print(f"Fehler: Die Datei {anhang_pfad} wurde nicht gefunden.")

    # Nachricht in Base64 codieren
    raw_message = base64.urlsafe_b64encode(email.as_bytes()).decode()
    return {'raw': raw_message}


def sende_email(dienst, empfaenger, betreff, nachricht, anhang_pfad):
    """Sendet die Email mit Gmail API."""
    email = erstelle_email_mit_anhang(empfaenger, betreff, nachricht, anhang_pfad)
    try:
        gesendete_email = dienst.users().messages().send(userId="me", body=email).execute()
        print(f"Email erfolgreich gesendet an: {empfaenger}. Nachrichten-ID: {gesendete_email['id']}")
    except Exception as e:
        print(f"Fehler beim Senden der Email: {e}")


def hauptprogramm():
    # Authentifizierung
    anmeldeinformationen = authentifizieren()
    dienst = build('gmail', 'v1', credentials=anmeldeinformationen)

    # Empfänger aus emails.txt lesen
    email_datei = "emails.txt"  # Datei mit Email-Adressen
    try:
        with open(email_datei, 'r', encoding='utf-8') as datei:
            empfaenger_liste = [zeile.strip() for zeile in datei if zeile.strip()]
    except FileNotFoundError:
        print(f"Fehler: Die Datei {email_datei} wurde nicht gefunden.")
        return

    # Nachricht aus message.txt lesen
    nachricht_datei = "message.txt"  # Datei mit Betreff und Nachricht
    try:
        with open(nachricht_datei, 'r', encoding='utf-8') as datei:
            zeilen = datei.readlines()
            betreff = zeilen[0].strip()  # Erste Zeile ist der Betreff
            nachricht = ''.join(zeilen[1:]).strip()  # Rest der Zeilen ist die Nachricht
    except FileNotFoundError:
        print(f"Fehler: Die Datei {nachricht_datei} wurde nicht gefunden.")
        return
    except IndexError:
        print(f"Fehler: Die Datei {nachricht_datei} ist leer oder falsch formatiert.")
        return

    # Pfad zur Anhang-Datei
    anhang_pfad = "cv.pdf"
    if not os.path.exists(anhang_pfad):
        print(f"Fehler: Die Anhang-Datei {anhang_pfad} wurde nicht gefunden.")
        return

    # Emails senden
    for empfaenger in empfaenger_liste:
        print(f"Email wird gesendet an: {empfaenger}")
        sende_email(dienst, empfaenger, betreff, nachricht, anhang_pfad)


if __name__ == "__main__":
    hauptprogramm()
