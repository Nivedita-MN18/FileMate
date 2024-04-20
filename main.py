import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QComboBox
from PIL import Image
import subprocess
from PyPDF2 import PdfFileReader, PdfFileWriter

class FileProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('File Processor')

        self.btn_select_input = QPushButton('Select Input File', self)
        self.btn_select_input.clicked.connect(self.select_input_file)

        self.btn_select_output = QPushButton('Select Output Folder', self)
        self.btn_select_output.clicked.connect(self.select_output_folder)

        self.operation_selector = QComboBox()
        self.operation_selector.addItem("Select Operation")
        self.operation_selector.addItem("Image to PDF")
        self.operation_selector.addItem("PDF to Image")
        self.operation_selector.addItem("Word to PDF")
        self.operation_selector.addItem("PDF to Word")
        self.operation_selector.addItem("Compress Image")
        self.operation_selector.addItem("Decompress Image")
        self.operation_selector.addItem("Convert Image Format")

        self.btn_process = QPushButton('Process', self)
        self.btn_process.clicked.connect(self.process_file)

        layout = QVBoxLayout()
        layout.addWidget(self.btn_select_input)
        layout.addWidget(self.btn_select_output)
        layout.addWidget(self.operation_selector)
        layout.addWidget(self.btn_process)

        self.setLayout(layout)

        self.input_file = None
        self.output_folder = None

    def select_input_file(self):
        options = QFileDialog.Options()
        self.input_file, _ = QFileDialog.getOpenFileName(self, "Select Input File", "", "All Files (*);;Image Files (*.jpg *.png *.bmp);;PDF Files (*.pdf);;Word Files (*.docx *.doc)", options=options)

    def select_output_folder(self):
        options = QFileDialog.Options()
        self.output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", options=options)

    def process_file(self):
        if self.input_file and self.output_folder:
            operation = self.operation_selector.currentText()

            if operation == "Image to PDF":
                self.img_to_pdf()
            elif operation == "PDF to Image":
                self.pdf_to_img()
            elif operation == "Word to PDF":
                self.word_to_pdf()
            elif operation == "PDF to Word":
                self.pdf_to_word()
            elif operation == "Compress Image":
                self.compress_image()
            elif operation == "Decompress Image":
                self.decompress_image()
            elif operation == "Convert Image Format":
                self.convert_image_format()
            else:
                print("Please select an operation.")

        else:
            print("Please select input file and output folder first.")

    def img_to_pdf(self):
        image = Image.open(self.input_file)
        output_path = os.path.join(self.output_folder, "output.pdf")
        image.convert("RGB").save(output_path, "PDF", resolution=100.0)
        print("Image to PDF conversion successful:", output_path)

    def pdf_to_img(self):
        output_path = os.path.join(self.output_folder, "output.png")
        subprocess.call(['pdftoppm', '-png', self.input_file, output_path])
        print("PDF to image conversion successful:", output_path)

    def word_to_pdf(self):
        output_path = os.path.join(self.output_folder, "output.pdf")
        subprocess.call(['libreoffice', '--headless', '--convert-to', 'pdf', self.input_file, '--outdir', self.output_folder])
        print("Word to PDF conversion successful:", output_path)

    def pdf_to_word(self):
        output_path = os.path.join(self.output_folder, "output.docx")
        subprocess.call(['pdftotext', self.input_file, output_path])
        print("PDF to Word conversion successful:", output_path)

    def compress_image(self):
        image = Image.open(self.input_file)
        output_path = os.path.join(self.output_folder, "compressed_" + os.path.basename(self.input_file))
        image.save(output_path, optimize=True, quality=50)  # Adjust quality as needed
        print("Image compression successful:", output_path)

    def decompress_image(self):
        image = Image.open(self.input_file)
        output_path = os.path.join(self.output_folder, "decompressed_" + os.path.basename(self.input_file))
        image.save(output_path)
        print("Image decompression successful:", output_path)

    def convert_image_format(self):
        image = Image.open(self.input_file)
        output_extension = ".jpg" if os.path.splitext(self.input_file)[1].lower() in ['.png', '.jpeg'] else ".png"
        output_path = os.path.join(self.output_folder, os.path.splitext(os.path.basename(self.input_file))[0] + output_extension)
        image.save(output_path)
        print("Image format conversion successful:", output_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileProcessor()
    window.show()
    sys.exit(app.exec_())
