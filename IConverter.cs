namespace HealthShare.PDFConverter
{
    public interface IConverter
    {
        bool ConvertToPDF(string inputFile, string outputFile);
    }
}
