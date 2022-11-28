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
