using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HealthShare.PDFConverter
{
    public class TIFFConverter : IConverter
    {
        public bool ConvertToPDF(string filename, string outfile)
        {
            bool converted = false;
            Document document = new Document();

            try
            {
                //load the tiff image and count the total pages  
                Bitmap bm = new Bitmap(filename);

                // creation of the different writers
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outfile, FileMode.Create));

                //Total number of pages
                int totalPages = bm.GetFrameCount(FrameDimension.Page);

                document.Open();

                PdfContentByte cb = writer.DirectContent;

                for (int pageNumber = 0; pageNumber < totalPages; ++pageNumber)
                {
                    bm.SelectActiveFrame(FrameDimension.Page, pageNumber);
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bm, ImageFormat.Jpeg);

                    // scale the image to fit in the page  
                    img.ScalePercent(7200f / img.DpiX, 7200f / img.DpiY);
                    img.SetAbsolutePosition(0, 0);

                    cb.AddImage(img);
                    document.NewPage();
                }
                converted = true;
                bm.Dispose();
            }

            catch
            {
                converted = false;
                throw;
            }

            finally
            {
                document.Close();
            }

            return converted;
        }

    }
}
