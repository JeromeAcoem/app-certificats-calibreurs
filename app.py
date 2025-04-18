import os
import re
import shutil
import zipfile
import tempfile
import pytesseract
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from collections import defaultdict, Counter
from datetime import datetime
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Certificats Calibreurs", layout="centered")
st.title("Certificats Calibreurs - Split & OCR")

uploaded_files = st.file_uploader("Chargez vos fichiers PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Lancer le traitement"):
        with st.spinner("Traitement en cours..."):
            temp_dir = tempfile.mkdtemp()
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)

            serial_number_pattern = re.compile(r"Serial number[:\s]+([\w\d-]+)", re.IGNORECASE)
            csv_data = []
            errors = []
            serial_tracker = defaultdict(int)
            total_files = 0

            for uploaded_file in uploaded_files:
                pdf_name = uploaded_file.name
                base_name = os.path.splitext(pdf_name)[0]
                pdf_path = os.path.join(temp_dir, pdf_name)

                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.read())

                try:
                    images = convert_from_path(pdf_path)
                except Exception as e:
                    errors.append(f"[LOAD ERROR] {pdf_name} : {str(e)}")
                    continue

                serial_numbers = []
                for i in range(0, len(images), 2):
                    try:
                        text = pytesseract.image_to_string(images[i])
                        match = serial_number_pattern.search(text)
                        serial = match.group(1) if match else f"Unknown_{i//2+1}"
                        serial_numbers.append(serial)
                    except Exception as e:
                        serial_numbers.append(f"Error_{i//2+1}")
                        errors.append(f"[OCR ERROR] {pdf_name} → Page {i+1} : {str(e)}")

                try:
                    reader = PdfReader(pdf_path)
                    num_pages = len(reader.pages)
                except Exception as e:
                    errors.append(f"[PDF READ ERROR] {pdf_name} : {str(e)}")
                    continue

                for i in range(0, num_pages, 2):
                    try:
                        writer = PdfWriter()
                        for j in range(2):
                            if i + j < num_pages:
                                writer.add_page(reader.pages[i + j])

                        serial_index = i // 2
                        serial_number = serial_numbers[serial_index] if serial_index < len(serial_numbers) else f"Unknown_{serial_index+1}"
                        serial_tracker[serial_number] += 1
                        is_duplicate = serial_tracker[serial_number] > 1

                        safe_serial = f"{serial_number}_{serial_tracker[serial_number]}" if is_duplicate else serial_number
                        output_filename = f"CAL31 - {base_name} - {safe_serial}.pdf"
                        output_path = os.path.join(output_dir, output_filename)

                        with open(output_path, "wb") as f_out:
                            writer.write(f_out)

                        csv_data.append({
                            "Fichier": output_filename,
                            "PDF source": base_name,
                            "Numéro de série": serial_number,
                            "Pages": f"{i+1}-{min(i+2, num_pages)}",
                            "Doublon": "oui" if is_duplicate else "non"
                        })

                        total_files += 1
                    except Exception as e:
                        errors.append(f"[PDF WRITE ERROR] {pdf_name} → Pages {i+1}-{i+2} : {str(e)}")

            # Rapport Excel
            df = pd.DataFrame(csv_data)
            excel_path = os.path.join(temp_dir, "rapport_certificats.xlsx")
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Certificats", index=False)
                if not df.empty:
                    summary = df.groupby("PDF source").agg({"Fichier": "count", "Doublon": lambda x: (x == "oui").sum()})
                    summary.rename(columns={"Fichier": "Total", "Doublon": "Doublons"}, inplace=True)
                    summary.to_excel(writer, sheet_name="Résumé")
                if errors:
                    pd.DataFrame(errors, columns=["Erreurs"]).to_excel(writer, sheet_name="Erreurs", index=False)

            # Archive ZIP
            zip_path = os.path.join(temp_dir, f"certificats_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.zip")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files_in_dir in os.walk(output_dir):
                    for file in files_in_dir:
                        full_path = os.path.join(root, file)
                        zipf.write(full_path, os.path.relpath(full_path, output_dir))
                zipf.write(excel_path, "rapport_certificats.xlsx")

            with open(zip_path, "rb") as f:
                st.success(f"Traitement terminé. {total_files} fichiers générés.")
                st.download_button("💾 Télécharger le ZIP", f.read(), file_name=os.path.basename(zip_path), mime="application/zip")
