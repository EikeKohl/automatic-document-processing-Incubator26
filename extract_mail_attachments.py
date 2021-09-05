from util.Extractor import Extractor
import os

MSG_DIR = "./Bestellungen"
SAVE_DIR = "./data"

Extractor.get_msg_file_attachments(msg_dir=MSG_DIR, save_dir=SAVE_DIR)

zip_files = [file for file in os.listdir(SAVE_DIR) if file.split(".")[-1] == "zip"]
pdf_files = [file for file in os.listdir(SAVE_DIR) if file.split(".")[-1] == "pdf"]

for zip_file in zip_files:
    file_path = os.path.join(SAVE_DIR, zip_file)
    Extractor.extract_zip(file_path, SAVE_DIR)

for pdf_file in pdf_files:
    file_path = os.path.join(SAVE_DIR, pdf_file)
    Extractor.convert_pdf_to_jpg(file_path)
