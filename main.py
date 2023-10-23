from spire.pdf.common import *
from spire.pdf import *

# Create a PdfDocument object
pdf = PdfDocument()
# Load a PDF document
pdf.LoadFromFile("Маршрутный Максимов.pdf")

# Create an XlsxLineLayoutOptions object to specify the conversion options
# Parameters: convertToMultipleSheet, rotatedText, splitCell, wrapText, overlapText
convertOptions = XlsxLineLayoutOptions(True, True, False, False, False)

# Set the conversion options
pdf.ConvertOptions.SetPdfToXlsxOptions(convertOptions)

# Save the PDF document to Excel XLSX format
pdf.SaveToFile("PdfToExcel_ОСИП.xlsx", FileFormat.XLSX)
pdf.Close()
print('jr')