using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Pdf;
using System;
using System.IO;

namespace HealthShare.PDFConverter
{
    public class WordConverter : IConverter
    {
        public bool ConvertToPDF(string inputFile, string outputFile)
        {
            bool converted = false;
            string extension = Path.GetExtension(inputFile);
            
            try
            {
                // Convert using SyncFusion.
                WordDocument wordDocument = new WordDocument();

                //Loads an existing Word document
                if (extension.ToUpper() == ".DOC")
                {
                    wordDocument = new WordDocument(inputFile, FormatType.Doc);
                }
                if (extension.ToUpper() == ".DOCX")
                {
                    wordDocument = new WordDocument(inputFile, FormatType.Docx);
                }
                if (extension.ToUpper() == ".DOT")
                {
                    wordDocument = new WordDocument(inputFile, FormatType.Dot);
                }

                //Initializes the ChartToImageConverter for converting charts during Word to pdf conversion
                wordDocument.ChartToImageConverter = new ChartToImageConverter();

                //Creates an instance of the DocToPDFConverter
                DocToPDFConverter converter = new DocToPDFConverter();

                //set PDF conformance level using DocToPDFConverterSettings class property.
                //converter.Settings.PdfConformanceLevel = PdfConformanceLevel.None;
                converter.Settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

                //Converts Word document into PDF document
                //PdfDocument pdfDocument = new PdfDocument(PdfConformanceLevel.Pdf_A1B);
                PdfDocument pdfDocument = new PdfDocument();
                pdfDocument = converter.ConvertToPDF(wordDocument);

                //Saves the PDF file 
                pdfDocument.Save(outputFile);

                //Closes the instance of document objects
                pdfDocument.Close(true);
                wordDocument.Close();

                converted = true;

            }
            catch (Exception)
            {
                
                throw;
            }

            return converted;
        }
    }
}
