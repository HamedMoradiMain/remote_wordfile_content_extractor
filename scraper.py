from zipfile import ZipFile 
from urllib.request import urlopen 
from io import BytesIO
from bs4 import BeautifulSoup
class WordReader:
    def wordfile(self):
        self.wordFile = urlopen("https://calibre-ebook.com/downloads/demos/demo.docx").read()
        self.wordFile= BytesIO(self.wordFile)
        self.document = ZipFile(self.wordFile)
        self.xml_content = self.document.read('word/document.xml')
        self.wordObj = BeautifulSoup(self.xml_content.decode('utf-8'),features="html.parser")
        self.textStrings = self.wordObj.findAll("w:t")
        for textElem in self.textStrings: 
            print(textElem.text)
    def run(self):
        print(self.wordfile())

if __name__ == "__main__":
    bot = WordReader()
    bot.run()





















'''from urllib.request import urlopen
from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO
from io import open '''

'''class PDFRader:
    PDFFile = urlopen("https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf")
    #PDFFile = open(r"F:\Books\Python\Black Hat Python, 2nd Edition by Justin Seitz  Tim Arnold [Justin Seitz].pdf",'rb')
    def readPDF(self,PDFFile):
        self.rsrmgr = PDFResourceManager()
        self.retstr = StringIO()
        self.laparams = LAParams()
        self.device = TextConverter(self.rsrmgr,self.retstr,laparams=self.laparams)
        process_pdf(self.rsrmgr,self.device,PDFFile)
        self.device.close()
        self.content = self.retstr.getvalue()
        self.retstr.close()
        return self.content
    def run(self):
        #print(self.PDFFile)
        outputString = self.readPDF(self.PDFFile)
        print(outputString)
        self.PDFFile.close()
if __name__ == "__main__":
    pdf = PDFRader()
    pdf.run()'''
