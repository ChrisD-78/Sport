#!/usr/bin/env python3
"""
Skript zum Extrahieren von Folien aus einer PowerPoint-Datei als Bilder
"""
import os
import sys

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    from PIL import Image
    import io
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

def extract_pptx_slides(pptx_path, output_dir):
    """Extrahiert Folien aus PowerPoint als Bilder"""
    if not HAS_PPTX:
        print("python-pptx nicht installiert. Installiere es...")
        os.system("python3 -m pip install python-pptx --quiet")
        try:
            from pptx import Presentation
        except ImportError:
            print("Fehler: Konnte python-pptx nicht installieren")
            return []
    
    os.makedirs(output_dir, exist_ok=True)
    
    prs = Presentation(pptx_path)
    slide_images = []
    
    print(f"Präsentation hat {len(prs.slides)} Folien")
    
    # Versuche, die Folien als Bilder zu exportieren
    # Da python-pptx keine direkte Bildkonvertierung unterstützt,
    # müssen wir einen anderen Ansatz verwenden
    
    # Alternative: Verwende LibreOffice oder unoconv falls verfügbar
    if os.system("which libreoffice > /dev/null 2>&1") == 0:
        print("Verwende LibreOffice zum Konvertieren...")
        cmd = f'libreoffice --headless --convert-to pdf --outdir "{output_dir}" "{pptx_path}"'
        os.system(cmd)
        pdf_path = os.path.join(output_dir, os.path.basename(pptx_path).replace('.pptx', '.pdf'))
        if os.path.exists(pdf_path):
            print(f"PDF erstellt: {pdf_path}")
            return extract_from_pdf(pdf_path, output_dir)
    
    # Fallback: Versuche mit unoconv
    if os.system("which unoconv > /dev/null 2>&1") == 0:
        print("Verwende unoconv zum Konvertieren...")
        cmd = f'unoconv -f pdf -o "{output_dir}" "{pptx_path}"'
        os.system(cmd)
        pdf_path = os.path.join(output_dir, os.path.basename(pptx_path).replace('.pptx', '.pdf'))
        if os.path.exists(pdf_path):
            print(f"PDF erstellt: {pdf_path}")
            return extract_from_pdf(pdf_path, output_dir)
    
    print("Hinweis: LibreOffice/unoconv nicht verfügbar. Verwende alternative Methode...")
    return []

def extract_from_pdf(pdf_path, output_dir):
    """Konvertiert PDF-Seiten zu Bildern"""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        images = []
        
        os.makedirs(os.path.join(output_dir, "slides"), exist_ok=True)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x Zoom für bessere Qualität
            image_path = os.path.join(output_dir, "slides", f"slide_{page_num + 1}.png")
            pix.save(image_path)
            images.append({
                'page': page_num + 1,
                'path': image_path,
                'filename': f"slide_{page_num + 1}.png"
            })
            print(f"Folie {page_num + 1} gespeichert: {image_path}")
        
        doc.close()
        return images
    except ImportError:
        print("PyMuPDF nicht verfügbar. Installiere es...")
        os.system("python3 -m pip install PyMuPDF --quiet")
        return extract_from_pdf(pdf_path, output_dir)
    except Exception as e:
        print(f"Fehler beim Extrahieren aus PDF: {e}")
        return []

def main():
    pptx_path = "/Users/christofdrost/Desktop/2025_Landau LaOla_Skizze Förderantrag.pptx"
    output_dir = "/Users/christofdrost/Sitzungsvorlage/pptx_content"
    
    if not os.path.exists(pptx_path):
        print(f"Fehler: PowerPoint-Datei nicht gefunden: {pptx_path}")
        return
    
    print(f"Analysiere PowerPoint: {pptx_path}")
    print(f"Ausgabe-Verzeichnis: {output_dir}")
    
    images = extract_pptx_slides(pptx_path, output_dir)
    
    if images:
        print(f"\n✓ {len(images)} Folien als Bilder extrahiert")
        print(f"Bilder gespeichert in: {output_dir}/slides/")
    else:
        print("\n⚠ Konnte keine Bilder extrahieren. Bitte installieren Sie LibreOffice oder verwenden Sie eine andere Methode.")

if __name__ == "__main__":
    main()


