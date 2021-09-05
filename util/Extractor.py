import pytesseract
import extract_msg
import os
import zipfile
import logging
from pdf2image import convert_from_path

"""
Extractor to extract email file attachments and texts from images.
"""

# logger = logging.getLogger('incubator')
logging.basicConfig(level=logging.INFO)


class Extractor:
    def __init__(self, img=None):
        self.img = img

    @classmethod
    def get_msg_file_attachments(cls, msg_dir, save_dir):
        """
        Extracts attachments of .msg files (Microsoft Outlook E-Mails) and
        saves them in a specified directory. Supported attachment types are
        .txt, .pdf, .jpeg, .png, .xml, .zip.

        Parameters
        ----------
        msg_dir : str
            directory containing msg files

        save_dir : str
            destination directory of attachments

        Returns
        -------
        None
        """
        allowed_data_formats = ["txt", "pdf", "jpeg", "png", "xml", "zip"]
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        msg_dir_content = os.listdir(msg_dir)
        files = [file for file in msg_dir_content if file.split(".")[-1] == "msg"]

        for file in range(len(files)):
            msg = extract_msg.Message(os.path.join(msg_dir, files[file]))
            for attachment in msg.attachments:
                file_name = attachment.longFilename
                data_format = file_name.split(".")[-1]
                if data_format in allowed_data_formats:
                    attachment.save(customPath=save_dir)
                    logging.info(f"file {file_name} has been extracted to {save_dir}")

    @classmethod
    def extract_zip(cls, zip_file, save_dir):
        """
        Extracts a zip_file and removes it afterwards.

        Parameters
        ----------
        zip_file : str
            filepath of zip directory to be unpacked
        save_dir : str
            target directory for zip file content

        Returns
        -------
        None
        """
        with zipfile.ZipFile(zip_file, "r") as zip_ref:
            zip_ref.extractall(save_dir)
            logging.info(f"zip file {zip_file} unpacked.")
        os.remove(zip_file)
        logging.info(f"zip file {zip_file} removed.")

    @classmethod
    def convert_pdf_to_jpg(cls, pdf):
        """
        Converts a pdf into JPEG on page level.

        Parameters
        ----------
        pdf : str
            filepath of pdf

        Returns
        -------
        None
        """

        img = convert_from_path(pdf, poppler_path = r"C:\Users\ekohl\OneDrive\Desktop\poppler-21.08.0\Library\bin")
        path, file_name = os.path.split(pdf)

        for i in range(len(img)):
            logging.info(f"converting {file_name}.")
            page_file_name = file_name + f"page_{i}.jpg", "JPEG"
            img[i].save(page_file_name)
            logging.info(f"{file_name} converted to {page_file_name}.")
            os.remove(pdf)
            logging.info(f"{file_name} deleted.")

    def extract_txt_from_img(self):
        text = pytesseract.image_to_string(self.img)
