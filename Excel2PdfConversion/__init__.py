__author__ = 'fahadadeel'
import jpype

class Excel2PdfConversion:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
    
    def main(self):
                
saveFormat = self.SaveFormat

workbook = self.Workbook(self.dataDir + "Book1.xls")

#Save the document in PDF format
workbook.save(self.dataDir + "OutBook1.pdf", saveFormat.PDF)

# Print message
print "\n Excel to PDF conversion performed successfully."
