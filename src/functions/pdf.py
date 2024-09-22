from PyPDF2 import PdfFileReader, PdfFileWriter
import io

def convert_pdf(pdf_file):
    pdf_reader = PdfFileReader(pdf_file)
    pdf_writer = PdfFileWriter()
    
    for page_num in range(pdf_reader.numPages):
        pdf_writer.addPage(pdf_reader.getPage(page_num))
    
    output = io.BytesIO()
    pdf_writer.write(output)
    output.seek(0)
    
    return output
