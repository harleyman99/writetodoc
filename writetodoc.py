import PySimpleGUI as sg
import os
from PyPDF2 import PdfReader
from docx import Document

def convert_pdf_to_docx(pdf_path, output_path):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        
        doc = Document()
        for page in pdf_reader.pages:
            text = page.extract_text()
            doc.add_paragraph(text)
        
        doc.save(output_path)

def main():
    layout = [
        [sg.Text("Select the PDF file:")],
        [sg.InputText(key='pdf_file_path'), sg.FileBrowse()],
        [sg.Button("Convert"), sg.Button("Exit")]
    ]

    window = sg.Window("PDF to Word Converter", layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "Exit":
            break
        elif event == "Convert":
            pdf_path = values['pdf_file_path']
            if pdf_path:
                output_path = os.path.splitext(pdf_path)[0] + ".docx"
                convert_pdf_to_docx(pdf_path, output_path)
                sg.popup("Conversion completed!", title="Success")
            else:
                sg.popup("Please select a PDF file.", title="Error")

    window.close()

if __name__ == "__main__":
    main()
