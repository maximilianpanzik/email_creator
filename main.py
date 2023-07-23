import pandas as pd
from outlook_msg import Message

# Funktion zum Ersetzen der Schlüsselworte in der Nachricht
def replace_keywords_in_message(msg_file, data_row):
    with open(msg_file, 'rb') as f:
        msg = Message(f)

    # Ersetze die Schlüsselworte im Nachrichtentext
    message_text = msg.plain_text
    for column_name, column_value in data_row.items():
        message_text = message_text.replace(f'{{{column_name}}}', column_value)

    msg.plain_text = message_text

    # Erstelle den Betreff und ersetze die Schlüsselworte
    subject = msg.subject
    for column_name, column_value in data_row.items():
        subject = subject.replace(f'{{{column_name}}}', column_value)

    msg.subject = subject

    # Speichere die bearbeitete Nachricht in einer neuen MSG-Datei
    new_msg_file = f'{data_row["Name"]}_Bestellnummer_{data_row["Bestellnummer"]}.msg'
    with open(new_msg_file, 'wb') as f:
        f.write(msg.as_bytes())

# Beispiel Excel-Daten (hier kannst du deine Excel-Daten einfügen)
data = [
    {'Name': 'Max Mustermann', 'Bestellnummer': 'ABC123', 'Produkt': 'Laptop', 'Preis': '999.99'},
    {'Name': 'Erika Musterfrau', 'Bestellnummer': 'XYZ789', 'Produkt': 'Smartphone', 'Preis': '499.99'}
]

# Pfad zur MSG-Vorlage und Bearbeitung der Nachrichten
msg_template_file = 'dein_template.msg'
for row in data:
    replace_keywords_in_message(msg_template_file, row)

print('E-Mail-Template in MSG-Dateien wurde erfolgreich erstellt!')

