"""
by Davide Rosa
Script per l'estrazione di immagini da PDF, esecuzione di OCR, traduzione e conversione in DOCX.
Questo script estrae immagini da file PDF, esegue l'OCR su di esse, traduce il testo estratto e lo converte in file DOCX.
Richiede le librerie PyMuPDF, Pillow, pytesseract, BeautifulSoup, docx e googletrans.
Per eseguire questo script, assicurati di avere installato le seguenti librerie:
- PyMuPDF (fitz)
- Pillow (PIL)
- pytesseract e https://github.com/tesseract-ocr/tesseract/releases
- BeautifulSoup
- docx
- googletrans
Questo script è progettato per essere eseguito in un ambiente Python con le librerie necessarie installate.
pip install -r requirements.txt
"""
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re
import os
from bs4 import BeautifulSoup
import shutil
import io  # Per gestire i byte delle immagini
from docx import Document  # Per creare file DOCX
from googletrans import Translator  # Per la traduzione automatica

def create_folder(folder):
    """Crea la cartella se non esiste."""
    if not os.path.exists(folder):
        os.makedirs(folder)

def clean_folder(folder):
    """Rimuove ricorsivamente il contenuto della cartella e la ricrea."""
    if os.path.exists(folder):
        shutil.rmtree(folder)
    create_folder(folder)

def estrai_immagini(pdf_path, output_folder):
    """
    Estrae tutte le immagini dal PDF, le ruota (se necessario) e le salva nella cartella output_folder.
    Il contatore delle immagini viene resettato per ogni PDF.
    """
    img_counter = 1
    print(f"Elaboro: {pdf_path}")
    documento = fitz.open(pdf_path)
    create_folder(output_folder)
    
    for i in range(len(documento)):
        pagina = documento.load_page(i)
        immagini = pagina.get_images(full=True)
        for img in immagini:
            xref = img[0]
            base_immagine = documento.extract_image(xref)
            immagine_bytes = base_immagine["image"]
            # Apri l'immagine da bytes
            image = Image.open(io.BytesIO(immagine_bytes))
            
            # Verifica se l'immagine necessita di essere ruotata per facilitare l'OCR
            try:
                osd = pytesseract.image_to_osd(image)
                match = re.search(r'Rotate: (\d+)', osd)
                if match:
                    rotation_angle = int(match.group(1))
                    if rotation_angle != 0:
                        print(f"Rotating image by {rotation_angle} degrees")
                        image = image.rotate(-rotation_angle, expand=True)
            except Exception as e:
                print("Errore durante il rilevamento dell'orientamento:", e)
            
            # Salva l'immagine (già ruotata se necessario)
            nome_immagine = os.path.join(output_folder, f"img_{img_counter:04d}.png")
            image.save(nome_immagine)
            img_counter += 1

def esegui_ocr_su_immagini(cartella_immagini, pdf_base, output_file_base, ocr_language='eng', translate_to=None):
    """
    Esegue l'OCR su tutte le immagini contenute in cartella_immagini e
    crea i seguenti file di output:
      - Nella cartella TXT/pdf_base/:
            • Un file di testo per ciascuna immagine
      - Nella cartella risultato/:
            • Un file aggregato (output_file_base) contenente tutto il testo estratto per questo PDF
      - Nella cartella HTML/pdf_base/:
            • File HTML e HOCR per ciascuna immagine, e un file aggregato HTML
    Restituisce il testo aggregato per questo PDF.
    """
    testo_completo = ""
    html_completo = ""
    
    # Crea la sottocartella per i file TXT individuali di questo PDF
    txt_pdf_folder = os.path.join("TXT", pdf_base)
    create_folder(txt_pdf_folder)
    
    # Crea la sottocartella per i file HTML di questo PDF
    html_pdf_folder = os.path.join("HTML", pdf_base)
    create_folder(html_pdf_folder)
    
    for immagine_nome in sorted(os.listdir(cartella_immagini)):
        if immagine_nome.lower().endswith('.png'):
            immagine_path = os.path.join(cartella_immagini, immagine_nome)
            print(f"Elaboro: {immagine_nome}")
            image = Image.open(immagine_path)
            
            # Esegui OCR sull'immagine
            extracted_text = pytesseract.image_to_string(image, lang=ocr_language)
            merged_text = re.sub(r'(\w+)-\n(\w+)', r'\1\2', extracted_text)
            testo_completo += merged_text + "\n"
            
            # Salva un file di testo per ciascuna immagine nella sottocartella TXT/pdf_base/
            with open(os.path.join(txt_pdf_folder, f"{immagine_nome}.txt"), "w", encoding="utf-8") as text_file:
                text_file.write(merged_text + "\n")
            
            # Genera l'output HOCR/HTML
            html_text = pytesseract.image_to_pdf_or_hocr(image, extension='hocr', lang=ocr_language)
            html_text_decoded = html_text.decode("utf-8")
            soup = BeautifulSoup(html_text_decoded, 'html.parser')
            formatted_text = soup.prettify()
            html_completo += formatted_text
            
            # Salva file HTML e HOCR per ciascuna immagine nella sottocartella HTML/pdf_base/
            with open(os.path.join(html_pdf_folder, f"{immagine_nome}.html"), "w", encoding="utf-8") as html_file:
                html_file.write(formatted_text)
            with open(os.path.join(html_pdf_folder, f"{immagine_nome}.hocr"), "w", encoding="utf-8") as hocr_file:
                hocr_file.write(html_text_decoded + "\n")
    
    # Salva il file TXT aggregato per questo PDF nella cartella "risultato"
    txt_output_path = os.path.join("risultato", output_file_base)
    with open(txt_output_path, "w", encoding="utf-8") as txt_out:
        txt_out.write(testo_completo)
        
    # Salva l'HTML aggregato in un file nella sottocartella HTML/pdf_base/
    html_output_path = os.path.join(html_pdf_folder, output_file_base + ".html")
    with open(html_output_path, "w", encoding="utf-8") as html_out:
        html_out.write(html_completo)
    
    # Se è specificata una lingua di traduzione, crea anche i file tradotti
    if translate_to and testo_completo.strip():
        print(f"Traduzione in corso verso {translate_to}...")
        testo_tradotto = translate_text(testo_completo, translate_to)
        
        # Salva il file TXT tradotto nella cartella "risultato"
        translated_filename = f"{pdf_base}_libro_translated_{translate_to}.txt"
        translated_txt_path = os.path.join("risultato", translated_filename)
        with open(translated_txt_path, "w", encoding="utf-8") as translated_out:
            translated_out.write(testo_tradotto)
        
        print(f"File tradotto salvato: {translated_filename}")
    
    return testo_completo

def translate_text(text, target_language):
    """
    Traduce il testo nella lingua target usando Google Translate.
    """
    translator = Translator()
    try:
        # Dividi il testo in chunks più piccoli per evitare errori di limite
        max_chunk_size = 4500  # Google Translate ha un limite di circa 5000 caratteri
        chunks = [text[i:i+max_chunk_size] for i in range(0, len(text), max_chunk_size)]
        
        translated_chunks = []
        for chunk in chunks:
            if chunk.strip():  # Solo se il chunk non è vuoto
                translated = translator.translate(chunk, dest=target_language)
                translated_chunks.append(translated.text)
            else:
                translated_chunks.append(chunk)
        
        return ''.join(translated_chunks)
    except Exception as e:
        print(f"Errore durante la traduzione: {e}")
        return text  # Ritorna il testo originale in caso di errore

def find_tesseract_executable():
    """
    Cerca automaticamente l'eseguibile di Tesseract in diverse posizioni comuni.
    Restituisce il percorso dell'eseguibile se trovato, altrimenti None.
    """
    import os
    
    # Possibili percorsi per Tesseract
    possible_paths = [
        # Percorsi standard di installazione
        r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe',
        r'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe',
        # Percorsi relativi alla cartella utente
        os.path.expanduser(r'~\AppData\\Local\\Tesseract-OCR\\tesseract.exe'),
        os.path.expanduser(r'~\AppData\\Roaming\\Tesseract-OCR\\tesseract.exe'),
        # Percorsi relativi alla cartella corrente
        r'.\\Tesseract-OCR\\tesseract.exe',
        r'.\\tesseract\\tesseract.exe',
        # Verifica se è già nel PATH
        'tesseract.exe',
        'tesseract'
    ]
    
    # Controlla ogni possibile percorso
    for path in possible_paths:
        try:
            # Per i percorsi assoluti, controlla se il file esiste
            if os.path.isabs(path) or path.startswith('.'):
                if os.path.isfile(path):
                    print(f"Tesseract trovato in: {path}")
                    return path
            else:
                # Per i comandi semplici, prova a eseguirli per vedere se sono nel PATH
                import subprocess
                try:
                    subprocess.run([path, '--version'], capture_output=True, check=True, timeout=5)
                    print(f"Tesseract trovato nel PATH: {path}")
                    return path
                except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
                    continue
        except Exception as e:
            continue
    
    print("ATTENZIONE: Tesseract non trovato automaticamente.")
    print("Assicurati che Tesseract sia installato e nel PATH, o modifica manualmente il percorso.")
    return None

def convert_txt_to_docx(txt_filepath, docx_filepath):
    """
    Converte un file TXT in un file DOCX.
    """
    document = Document()
    with open(txt_filepath, "r", encoding="utf-8") as f:
        content = f.read()
    document.add_paragraph(content)
    document.save(docx_filepath)

if __name__ == '__main__':
    pdf_folder = "pdf"           # Cartella contenente i file PDF
    img_root_folder = "IMG"      # Cartella principale per le immagini estratte
    
    # Trova automaticamente il percorso di Tesseract
    tesseract_path = find_tesseract_executable()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    else:
        # Fallback al percorso predefinito se non trovato automaticamente
        pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
        print("Utilizzo il percorso predefinito di Tesseract. Se non funziona, installa Tesseract o modifica il percorso.")
    
    # Chiede all'utente di scegliere la lingua per l'OCR
    print("Scegli la lingua per l'OCR:")
    print("1. Italiano (ita)")
    print("2. Inglese (eng)")
    
    while True:
        scelta = input("Inserisci 1 o 2: ").strip()
        if scelta == "1":
            ocr_language = "ita"
            translate_language = "en"  # Traduce in inglese
            print("Lingua selezionata: Italiano")
            print("I testi verranno anche tradotti in inglese")
            break
        elif scelta == "2":
            ocr_language = "eng"
            translate_language = "it"  # Traduce in italiano
            print("Lingua selezionata: Inglese")
            print("I testi verranno anche tradotti in italiano")
            break
        else:
            print("Scelta non valida. Inserisci 1 per Italiano o 2 per Inglese.")
    
    # Pulisce le cartelle di output principali
    clean_folder(img_root_folder)
    clean_folder("TXT")
    clean_folder("HTML")
    clean_folder("risultato")
    
    # Variabile per accumulare il testo di tutti i PDF
    global_aggregated_text = ""
    
    # Itera su tutti i PDF nella cartella 'pdf'
    for pdf_file in sorted(os.listdir(pdf_folder)):
        if pdf_file.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            pdf_base = os.path.splitext(pdf_file)[0]
            
            # Crea una sottocartella per le immagini estratte da questo PDF
            pdf_img_folder = os.path.join(img_root_folder, pdf_base)
            estrai_immagini(pdf_path, pdf_img_folder)
            
            # Definisce il nome del file di output aggregato per il PDF corrente
            output_txt_file = f"{pdf_base}_libro.txt"
            testo_pdf = esegui_ocr_su_immagini(pdf_img_folder, pdf_base, output_txt_file, ocr_language, translate_language)
            
            # Aggiunge al testo globale una separazione per ogni PDF
            global_aggregated_text += f"\n--- {pdf_base} ---\n" + testo_pdf
    
    # Salva il file aggregato globale "libro.txt" nella cartella "risultato"
    global_output_path = os.path.join("risultato", "libro.txt")
    with open(global_output_path, "w", encoding="utf-8") as global_out:
        global_out.write(global_aggregated_text)
    
    # Crea anche la traduzione del file aggregato globale
    if global_aggregated_text.strip():
        print(f"Traduzione del file aggregato globale in corso verso {translate_language}...")
        global_translated_text = translate_text(global_aggregated_text, translate_language)
        
        global_translated_path = os.path.join("risultato", f"libro_translated_{translate_language}.txt")
        with open(global_translated_path, "w", encoding="utf-8") as global_translated_out:
            global_translated_out.write(global_translated_text)
        
        print(f"File aggregato tradotto salvato: libro_translated_{translate_language}.txt")
    
    # Converte ogni file TXT in "risultato" in un file DOCX
    for filename in os.listdir("risultato"):
        if filename.lower().endswith(".txt"):
            txt_filepath = os.path.join("risultato", filename)
            docx_filepath = os.path.join("risultato", os.path.splitext(filename)[0] + ".docx")
            print(f"Converting {txt_filepath} to {docx_filepath}")
            convert_txt_to_docx(txt_filepath, docx_filepath)
