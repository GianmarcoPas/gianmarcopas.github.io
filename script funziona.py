import openpyxl
from bs4 import BeautifulSoup

# Percorso del file Excel
excel_file_path = r"C:\Users\passi\Desktop\Lavoro\Pianificazione.xlsx"

# Percorso del file HTML
html_file_path = r"C:\Users\passi\gianmarcopas.github.io\index.html"

# Carica il file Excel
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Apri il file HTML
with open(html_file_path, 'r', encoding='utf-8') as html_file:
    html_content = html_file.read()

# Analizza il file HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Trova la tabella nel file HTML
table = soup.find('table')

# Ottieni le righe della tabella HTML
rows = table.find_all('tr')

# Rimuovi tutte le righe tranne l'intestazione
for row in rows[1:]:
    row.extract()

# Itera attraverso le righe del foglio Excel e crea nuove righe nel file HTML
for excel_row in sheet.iter_rows(min_row=2):
    new_row = soup.new_tag('tr')
    
    # Estrai i valori dal foglio Excel
    values = [cell.value if cell.value is not None else "" for cell in excel_row]
    
    # Crea nuove celle nella riga HTML
    for value in values:
        cell = soup.new_tag('td')
        cell.string = str(value)
        new_row.append(cell)
    
    # Aggiungi la nuova riga alla tabella HTML
    table.append(new_row)

# Sovrascrivi il file HTML con le nuove righe
with open(html_file_path, 'w', encoding='utf-8') as html_file:
    html_file.write(str(soup))

print("Modifiche apportate con successo.")

# Chiudi il file Excel
workbook.close()
