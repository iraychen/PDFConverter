using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;
using System.IO;

namespace HealthShare.PDFConverter
{
    public class ImageConverter : IConverter
    {
        public bool ConvertToPDF(string inputFile, string outputFile)
        {
            bool converted = false;
            iTextSharp.text.Rectangle pageSize = null;

            using (var srcImage = new Bitmap(inputFile))
            {
                pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
            }

            using (var ms = new MemoryStream())
            {
                Document document = new Document(pageSize, 0, 0, 0, 0);
                try
                {
                    PdfWriter.GetInstance(document, ms).SetFullCompression();

                    document.Open();

                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputFile);
                    document.Add(image);
                    document.Close();
                    File.WriteAllBytes(outputFile, ms.ToArray());
                    converted = true;
                }
                catch
                {
                    converted = false;
                    throw;
                }
                finally
                {
                }
            }

            return converted;
        }
    }
}
