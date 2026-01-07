#!/usr/bin/env python3
"""
Skript zum Extrahieren von Bildern und Inhalten aus der PDF-Datei
"""
import os
import sys

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    from pdf2image import convert_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

def extract_with_pymupdf(pdf_path, output_dir):
    """Extrahiert Bilder und Text mit PyMuPDF"""
    doc = fitz.open(pdf_path)
    images = []
    text_content = []
    
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Text extrahieren
        text = page.get_text()
        text_content.append({
            'page': page_num + 1,
            'text': text
        })
        
        # Bilder extrahieren
        image_list = page.get_images()
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            image_filename = f"page_{page_num + 1}_img_{img_index + 1}.{image_ext}"
            image_path = os.path.join(output_dir, "images", image_filename)
            
            with open(image_path, "wb") as img_file:
                img_file.write(image_bytes)
            
            images.append({
                'page': page_num + 1,
                'filename': image_filename,
                'path': image_path
            })
    
    doc.close()
    return images, text_content

def extract_with_pdf2image(pdf_path, output_dir):
    """Konvertiert PDF-Seiten zu Bildern mit pdf2image"""
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, "page_images"), exist_ok=True)
    
    images = convert_from_path(pdf_path, dpi=200)
    image_paths = []
    
    for i, image in enumerate(images):
        image_filename = f"page_{i + 1}.png"
        image_path = os.path.join(output_dir, "page_images", image_filename)
        image.save(image_path, "PNG")
        image_paths.append(image_path)
    
    return image_paths

def extract_text_with_pdfplumber(pdf_path):
    """Extrahiert Text mit pdfplumber"""
    text_content = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            text_content.append({
                'page': page_num + 1,
                'text': text or ''
            })
    return text_content

def main():
    pdf_path = "/Users/christofdrost/Desktop/Sitzungsvorlage.pdf"
    output_dir = "/Users/christofdrost/Sitzungsvorlage/pdf_content"
    
    if not os.path.exists(pdf_path):
        print(f"Fehler: PDF-Datei nicht gefunden: {pdf_path}")
        return
    
    print(f"Analysiere PDF: {pdf_path}")
    print(f"Ausgabe-Verzeichnis: {output_dir}")
    
    images = []
    text_content = []
    
    # Versuche PyMuPDF zuerst (beste Qualität für Bilder)
    if HAS_PYMUPDF:
        print("\nVerwende PyMuPDF zum Extrahieren...")
        images, text_content = extract_with_pymupdf(pdf_path, output_dir)
        print(f"✓ {len(images)} Bilder extrahiert")
        print(f"✓ {len(text_content)} Seiten Text extrahiert")
    elif HAS_PDF2IMAGE:
        print("\nVerwende pdf2image zum Konvertieren...")
        image_paths = extract_with_pdf2image(pdf_path, output_dir)
        print(f"✓ {len(image_paths)} Seiten als Bilder gespeichert")
    
    # Text mit pdfplumber extrahieren (falls verfügbar)
    if HAS_PDFPLUMBER and not text_content:
        print("\nExtrahiere Text mit pdfplumber...")
        text_content = extract_text_with_pdfplumber(pdf_path)
        print(f"✓ {len(text_content)} Seiten Text extrahiert")
    
    # Speichere Text-Inhalte
    if text_content:
        text_file = os.path.join(output_dir, "extracted_text.txt")
        with open(text_file, "w", encoding="utf-8") as f:
            for page_data in text_content:
                f.write(f"\n{'='*60}\n")
                f.write(f"SEITE {page_data['page']}\n")
                f.write(f"{'='*60}\n\n")
                f.write(page_data['text'])
                f.write("\n\n")
        print(f"✓ Text gespeichert in: {text_file}")
    
    # Stelle sicher, dass das Ausgabeverzeichnis existiert
    os.makedirs(output_dir, exist_ok=True)
    
    # Erstelle Übersicht
    overview_file = os.path.join(output_dir, "overview.txt")
    with open(overview_file, "w", encoding="utf-8") as f:
        f.write("PDF-INHALT ÜBERSICHT\n")
        f.write("="*60 + "\n\n")
        f.write(f"Anzahl Seiten: {len(text_content) if text_content else 'Unbekannt'}\n")
        f.write(f"Anzahl Bilder: {len(images)}\n\n")
        
        if images:
            f.write("EXTRAHIERTE BILDER:\n")
            f.write("-"*60 + "\n")
            for img in images:
                f.write(f"Seite {img['page']}: {img['filename']}\n")
            f.write("\n")
        
        if text_content:
            f.write("\nTEXT-INHALTE (Auszug):\n")
            f.write("-"*60 + "\n")
            for page_data in text_content[:3]:  # Erste 3 Seiten als Beispiel
                f.write(f"\nSeite {page_data['page']}:\n")
                preview = page_data['text'][:200].replace('\n', ' ')
                f.write(f"{preview}...\n")
    
    print(f"\n✓ Übersicht gespeichert in: {overview_file}")
    print(f"\nFertig! Alle Inhalte wurden nach '{output_dir}' extrahiert.")

if __name__ == "__main__":
    main()

