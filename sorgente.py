import os
import glob
import subprocess
import platform
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter
from datetime import datetime as dt

def estrai_xml_da_p7m(p7m_path, xml_output_path):
    """Estrae l'XML da un file .p7m usando openssl"""
    try:
        cmd = [
            "openssl", "smime", "-verify", "-noverify",
            "-in", p7m_path,
            "-inform", "DER",
            "-out", xml_output_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        try:
            cmd[cmd.index("DER")] = "PEM"
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception as e:
            print(f"Errore nell'estrazione del p7m {os.path.basename(p7m_path)}: {e}")
            return False

def get_tag_text(element, path):
    parts = path.split('/')
    current = element
    for part in parts:
        found = False
        for child in current:
            tag_name = child.tag.split('}')[-1]
            if tag_name == part:
                current = child
                found = True
                break
        if not found:
            return ""
    return current.text if current.text else ""

def xml_to_single_pdf(xml_file, pdf_output_path):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    doc = SimpleDocTemplate(
        pdf_output_path, 
        pagesize=letter,
        rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
    )
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=14, leading=18, textColor=colors.HexColor("#1A365D"))
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=9, leading=12)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8.5, leading=11)
    cell_bold = ParagraphStyle('CellBold', parent=cell_style, fontName='Helvetica-Bold')
    
    cedente_denominazione = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione')
    ced_indirizzo = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/Sede/Indirizzo')
    ced_cap = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/Sede/CAP')
    ced_comune = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/Sede/Comune')
    ced_prov = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/Sede/Provincia')
    ced_piva = get_tag_text(root, 'FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice')
    
    cessionario_denominazione = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione')
    ces_indirizzo = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/Sede/Indirizzo')
    ces_cap = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/Sede/CAP')
    ces_comune = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/Sede/Comune')
    ces_prov = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/Sede/Provincia')
    ces_cf = get_tag_text(root, 'FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/CodiceFiscale')

    num_fattura = get_tag_text(root, 'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero')
    data_fattura = get_tag_text(root, 'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data')
    divisa = get_tag_text(root, 'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Divisa')
    
    story.append(Paragraph(f"<b>FATTURA ELETTRONICA N. {num_fattura}</b> del {data_fattura}", title_style))
    story.append(Spacer(1, 10))
    
    cedente_text = f"<b>CEDENTE / PRESTATORE</b><br/>{cedente_denominazione}<br/>{ced_indirizzo}<br/>{ced_cap} {ced_comune} ({ced_prov})<br/>P.IVA: {ced_piva}"
    cessionario_text = f"<b>CESSIONARIO / COMMITTENTE</b><br/>{cessionario_denominazione}<br/>{ces_indirizzo}<br/>{ces_cap} {ces_comune} ({ces_prov})<br/>C.F.: {ces_cf}"
    
    parties_table = Table([[Paragraph(cedente_text, header_style), Paragraph(cessionario_text, header_style)]], colWidths=[270, 270])
    parties_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(parties_table)
    story.append(Spacer(1, 10))
    
    table_data = [[
        Paragraph("<b>N°</b>", cell_bold),
        Paragraph("<b>Descrizione</b>", cell_bold),
        Paragraph("<b>Q.tà</b>", cell_bold),
        Paragraph("<b>Prezzo Unit.</b>", cell_bold),
        Paragraph("<b>IVA</b>", cell_bold),
        Paragraph("<b>Prezzo Totale</b>", cell_bold)
    ]]
    
    for body in root.findall('.//{*}FatturaElettronicaBody'):
        for linea in body.findall('.//{*}DettaglioLinee'):
            num_linea = get_tag_text(linea, 'NumeroLinea')
            desc = get_tag_text(linea, 'Descrizione')
            qta = get_tag_text(linea, 'Quantita')
            prezzo_unit = get_tag_text(linea, 'PrezzoUnitario')
            prezzo_tot = get_tag_text(linea, 'PrezzoTotale')
            aliquota = get_tag_text(linea, 'AliquotaIVA')
            
            table_data.append([
                Paragraph(num_linea, cell_style),
                Paragraph(desc, cell_style),
                Paragraph(qta, cell_style),
                Paragraph(f"{float(prezzo_unit):.2f}", cell_style),
                Paragraph(f"{float(aliquota):.1f}%", cell_style),
                Paragraph(f"{float(prezzo_tot):.2f}", cell_style)
            ])
            
    items_table = Table(table_data, colWidths=[30, 240, 45, 65, 50, 70])
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(items_table)
    story.append(Spacer(1, 10))
    
    imponibile_tot = 0
    imposta_tot = 0
    for body in root.findall('.//{*}FatturaElettronicaBody'):
        for riepilogo in body.findall('.//{*}DatiRiepilogo'):
            imponibile_tot += float(get_tag_text(riepilogo, 'ImponibileImporto') or 0)
            imposta_tot += float(get_tag_text(riepilogo, 'Imposta') or 0)
            
    totale_documento = imponibile_tot + imposta_tot
    
    summary_data = [
        [Paragraph("<b>Imponibile:</b>", cell_style), Paragraph(f"{imponibile_tot:.2f} {divisa}", cell_style)],
        [Paragraph("<b>Imposta:</b>", cell_style), Paragraph(f"{imposta_tot:.2f} {divisa}", cell_style)],
        [Paragraph("<b>Totale Documento:</b>", cell_bold), Paragraph(f"<b>{totale_documento:.2f} {divisa}</b>", cell_bold)],
    ]
    
    summary_table = Table(summary_data, colWidths=[110, 90])
    summary_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('LINEABOVE', (0,2), (-1,2), 1, colors.HexColor("#1A365D")),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    
    wrapper_table = Table([["", summary_table]], colWidths=[340, 200])
    wrapper_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(wrapper_table)
    
    doc.build(story)

def elabora_cartella(cartella_input):
    now = (dt.now()).strftime("%d-%m-%Y")
    pdf_finale = os.path.join(cartella_input, f"fatture_{now}.pdf")
    
    os.makedirs("temp_pdf", exist_ok=True)
    os.makedirs("temp_xml", exist_ok=True)
    writer = PdfWriter()
    
    files_input = glob.glob(os.path.join(cartella_input, "*.xml")) + glob.glob(os.path.join(cartella_input, "*.p7m"))
    if not files_input:
        return None, "Nessun file XML o P7M trovato nella cartella selezionata!"

    print(f"{len(files_input)} file trovati. Avvio elaborazione...")

    for i, file_path in enumerate(files_input):
        estensione = os.path.splitext(file_path)[1].lower()
        xml_path = file_path
        
        if estensione == ".p7m":
            xml_path = f"temp_xml/estratto_{i}.xml"
            successo = estrai_xml_da_p7m(file_path, xml_path)
            if not successo:
                continue
        
        temp_pdf = f"temp_pdf/temp_{i}.pdf"
        try:
            xml_to_single_pdf(xml_path, temp_pdf)
            reader = PdfReader(temp_pdf)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"Errore nella generazione del PDF per {os.path.basename(file_path)}: {e}")

    with open(pdf_finale, "wb") as output_file:
        writer.write(output_file)
    
    for f in glob.glob("temp_pdf/*.pdf") + glob.glob("temp_xml/*.xml"):
        os.remove(f)
    os.rmdir("temp_pdf")
    os.rmdir("temp_xml")
    
    print(f"\nFatto! Tutte le fatture sono state convertite in: '{pdf_finale}'.")
    return pdf_finale, None

class InterfacciaFatture:
    def __init__(self, root):
        self.root = root
        self.root.title("Hoppydf 26.10 - Fatture XML/P7M")
        self.root.geometry("460x220")
        self.root.resizable(False, False)
        
        self.cartella_selezionata = tk.StringVar()
        
        tk.Label(root, text="Benvenuto!\nPer iniziare seleziona la cartella di origine delle fatture:", font=("Helvetica", 10, "bold")).pack(pady=15)
        
        frame_input = tk.Frame(root)
        frame_input.pack(pady=5, padx=20, fill="x")
        
        self.entry_path = tk.Entry(frame_input, textvariable=self.cartella_selezionata, state="readonly", width=42, font=("Helvetica", 9))
        self.entry_path.pack(side="left", padx=(0, 10))
        
        btn_sfoglia = tk.Button(frame_input, text="Sfoglia...", command=self.seleziona_cartella, font=("Helvetica", 9))
        btn_sfoglia.pack(side="left")
        
        frame_bottoni = tk.Frame(root)
        frame_bottoni.pack(pady=25)
        
        self.btn_avvia = tk.Button(frame_bottoni, text="Avvia conversione", font=("Helvetica", 10, "bold"), bg="#1A365D", fg="white", padx=10, pady=6, command=self.esegui_conversione, state="disabled")
        self.btn_avvia.pack(side="left", padx=12)
        
        btn_chiudi = tk.Button(frame_bottoni, text="Chiudi", font=("Helvetica", 10, "bold"), bg="#A00", fg="white", padx=18, pady=6, command=root.destroy)
        btn_chiudi.pack(side="left", padx=12)

    def seleziona_cartella(self):
        cartella = filedialog.askdirectory(title="Seleziona cartella di origine")
        if cartella:
            self.cartella_selezionata.set(cartella)
            self.btn_avvia.config(state="normal")

    def esegui_conversione(self):
        cartella = self.cartella_selezionata.get()
        if not cartella:
            messagebox.showwarning("Attenzione", "Seleziona prima una cartella di origine!")
            return
        
        try:
            pdf_path, errore = elabora_cartella(cartella)
            if errore:
                messagebox.showerror("Errore", errore)
                return
            
            apri_file = messagebox.askyesno(
                "Completato", 
                f"Conversione completata con successo!\n\nIl file è stato salvato in:\n{pdf_path}\n\nDesideri aprire il file PDF adesso?"
            )
            if apri_file:
                self.apri_pdf(pdf_path)
                
        except Exception as e:
            messagebox.showerror("Errore Critico", f"Si è verificato un errore inatteso:\n{e}")

    def apri_pdf(self, path):
        try:
            if platform.system() == "Darwin":
                subprocess.run(["open", path])
            elif platform.system() == "Windows":
                os.startfile(path)
            else:
                subprocess.run(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire automaticamente il file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfacciaFatture(root)
    root.mainloop()
