import os
import tempfile
import subprocess
from pathlib import Path
from io import BytesIO
import comtypes.client
from pptx import Presentation
import requests


class PDFConverter:
    @staticmethod
    def pptx_to_pdf_linux(pptx_path: str, pdf_path: str) -> bool:
        try:
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(Path(pdf_path).parent),
                pptx_path,
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            return result.returncode == 0
        except Exception as e:
            print(f"Error converting PPTX to PDF: {e}")
            return False

        return True

    @staticmethod
    def pptx_to_pdf_windows(pptx_path: str, pdf_path: str) -> bool:
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1

            deck = powerpoint.Presentations.Open(pptx_path)
            deck.SaveAs(pdf_path, FileFormat=32)
            deck.Close()
            powerpoint.Quit()
            return True
        except Exception as e:
            print(f"Error converting PPTX to PDF: {e}")
            return False

    @staticmethod
    def convert_pptx_to_pdf(pptx_bytes: BytesIO) -> BytesIO:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as pptx_temp:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as pdf_temp:
                try:
                    pptx_temp.write(pptx_bytes.getbuffer())
                    pptx_temp.flush()

                    if os.name == "nt":
                        success = PDFConverter.pptx_to_pdf_windows(
                            pptx_temp.name, pdf_temp.name
                        )
                    else:
                        success = PDFConverter.pptx_to_pdf_linux(
                            pptx_temp.name, pdf_temp.name
                        )

                    if success and os.path.exists(pdf_temp.name):
                        with open(pdf_temp.name, "rb") as f:
                            pdf_bytes = BytesIO(f.read())
                        return pdf_bytes
                    else:
                        raise Exception("Failed to convert PPTX to PDF.")
                finally:
                    for temp_file in [pptx_temp.name, pdf_temp.name]:
                        if os.path.exists(temp_file):
                            try:
                                os.unlink(temp_file)
                            except:
                                pass

class OnlinePDFConverter:
    @staticmethod
    def convert_pptx_to_pdf(pptx_bytes: BytesIO) -> BytesIO:
        try:
            url="https://api.cloudconvert.com/v2/convert"
            files = {
                'file': ('silicon_eic_template.pptx', pptx_bytes.getvalue(), 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            }
            data = {
                "api_key": "YOUR_API_KEY",
            }
            response = requests.post(url, files=files, data=data, timeout=120)
            if response.status_code == 200:
                return BytesIO(response.content)
            
            else:
                raise Exception(f"Failed to convert PPTX to PDF. Status code: {response.status_code}")
        except Exception as e:
            print(f"Error converting PPTX to PDF using online service: {e}")
            raise
