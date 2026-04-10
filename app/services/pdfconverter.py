import io
import os
import platform
import shutil
import subprocess
import tempfile
import threading
from typing import Optional


class PDFConverter:
    _convert_fn = None
    _convert_lock = threading.Lock()

    @classmethod
    def _get_converter(cls):
        if cls._convert_fn is not None:
            return cls._convert_fn

        with cls._convert_lock:
            if cls._convert_fn is None:
                try:
                    from docx2pdf import convert as docx2pdf_convert
                except ImportError as import_error:
                    raise Exception(
                        "docx2pdf library is not installed. "
                        "Install it using: pip install docx2pdf"
                    ) from import_error
                cls._convert_fn = docx2pdf_convert

        return cls._convert_fn

    @staticmethod
    def _convert_with_docx2pdf(docx_path: str, pdf_path: str) -> None:
        docx2pdf_convert = PDFConverter._get_converter()

        try:
            docx2pdf_convert(docx_path, pdf_path, keep_active=True)
        except TypeError:
            docx2pdf_convert(docx_path, pdf_path)

    @staticmethod
    def _convert_with_libreoffice(docx_path: str, output_dir: str) -> str:
        binary = shutil.which("soffice") or shutil.which("libreoffice")
        if not binary:
            raise Exception(
                "LibreOffice is required on Linux for DOCX to PDF conversion. "
                "Install it with: apt-get update && apt-get install -y libreoffice"
            )

        result = subprocess.run(
            [
                binary,
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                output_dir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            check=False,
        )

        generated_pdf_path = os.path.join(
            output_dir,
            f"{os.path.splitext(os.path.basename(docx_path))[0]}.pdf",
        )

        if result.returncode != 0 or not os.path.exists(generated_pdf_path):
            stderr = (result.stderr or "").strip()
            stdout = (result.stdout or "").strip()
            details = stderr or stdout or "No converter output"
            raise Exception(f"LibreOffice conversion failed: {details}")

        return generated_pdf_path


    @staticmethod
    def convert_docx_to_pdf(docx_stream: io.BytesIO, output_file_name: Optional[str] = None) -> io.BytesIO:

        temp_dir = None
        try:
            temp_dir = tempfile.mkdtemp()
            
            docx_path = os.path.join(temp_dir, "input.docx")
            docx_stream.seek(0)
            with open(docx_path, "wb") as f:
                f.write(docx_stream.read())
            
            pdf_path = os.path.join(temp_dir, "output.pdf")

            if platform.system().lower() == "linux":
                generated_pdf_path = PDFConverter._convert_with_libreoffice(docx_path, temp_dir)
                if generated_pdf_path != pdf_path:
                    shutil.copyfile(generated_pdf_path, pdf_path)
            else:
                PDFConverter._convert_with_docx2pdf(docx_path, pdf_path)
            
            if not os.path.exists(pdf_path):
                raise Exception("PDF file was not created during conversion")
            
            with open(pdf_path, "rb") as f:
                pdf_stream = io.BytesIO(f.read())
            
            pdf_stream.seek(0)
            return pdf_stream

        except Exception as e:
            raise Exception(f"PDF conversion failed: {str(e)}")
        finally:
            if temp_dir and os.path.exists(temp_dir):
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass
