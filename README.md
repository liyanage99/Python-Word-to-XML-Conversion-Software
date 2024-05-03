Word to XML Converter
This is a Python application that converts .docx files into XML format using specific rules for text extraction and formatting. The converted XML files can be used for further processing or analysis.

Features
Extracts text and formatting (bold, italic) from Word documents.
Detects language of the document and segments content accordingly.
Converts extracted content into structured XML format.

Requirements
Python 3.x
Libraries: tkinter, docx, xml.etree.ElementTree, langdetect, datefinder

Installation
Clone this repository to your local machine:
bash
Copy code
git clone https://github.com/your-username/word-to-xml-converter.git
Navigate to the project directory:
bash
Copy code
cd word-to-xml-converter
Install the required libraries:
bash
Copy code
pip install -r requirements.txt
Usage
Ensure you have Python and the required libraries installed.
Run the application by executing main.py:
bash
Copy code
python main.py
The GUI will appear. Follow these steps:
Select a Folder: Click "Browse" to choose the folder containing your .docx files.
Enter the Starting Number of QR Code: Input the starting number for QR Code generation.
Click "Convert": Begin the conversion process.

Output

Converted XML files will be saved in an output folder within the selected folder.

Notes

Only .docx files are supported.
Ensure that the selected folder contains .docx files for conversion.
The starting number for QR Code is used to generate unique identifiers in the XML output.

Example
Suppose you have a folder named input_files containing .docx files. After conversion, XML files will be generated in the input_files/output folder.

Author
Liyanagecmadusanka2018@gmail.com
