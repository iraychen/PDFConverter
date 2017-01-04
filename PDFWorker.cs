using System;
using System.IO;

namespace HealthShare.PDFConverter
{
    public class PDFWorker
    {
        // File to be converted
        public string InputFile { get; set; }

        // Output pdf file.
        public string OutputFile { get; set; }

        // Mime type
        public string MimeType { get; set; }

        // Any identifier
        public string ReferenceId { get; set; }

        // Output message upon conversion.
        public string Message { get; set; }

        /// <summary>
        /// Constructor.
        /// Sets the properties.
        /// </summary>
        /// <param name="input">File to be converted</param>
        /// <param name="mimeType">Mime type of the input file</param>
        /// <param name="output">Output pdf file name</param>
        /// <param name="referenceId">Any identifier for the currently processed file</param>
        public PDFWorker(string input, string mimeType, string output, string referenceId = "")
        {
            InputFile = input;
            OutputFile = output;
            MimeType = mimeType;
            ReferenceId = referenceId;
        }

        /// <summary>
        /// Checks if the mimetype is supported or not.
        /// </summary>
        /// <returns>Returns true if the mimetype is supported</returns>
        public bool IsSupportedMimeType()
        {
            bool isSupported = false;

            switch (MimeType)
            {
                case Constant.MimeType.IMAGE_GIF:
                case Constant.MimeType.IMAGE_JPG:
                case Constant.MimeType.IMAGE_PJPEG:
                case Constant.MimeType.IMAGE_TIF:
                case Constant.MimeType.IMAGE_TIFF:
                case Constant.MimeType.TEXT_HTML:
                case Constant.MimeType.APPLICATION_MSWORD:
                case Constant.MimeType.APPLICATION_OCTET_STREAM:
                case Constant.MimeType.APPLICATION_OPENXMLFORMATS_OFFICEDOCUMENT:
                case Constant.MimeType.APPLICATION_PDF:
                    {
                        // PDFs are not converted.
                        isSupported = true;
                        break;
                    }
                default: // Unsupported mimetype which cannot be converted.
                    {
                        isSupported = false;
                        break;
                    }
            }

            return isSupported;
        }

        /// <summary>
        /// Converts the input file into pdf file based on the mimetype.
        /// Supported mime types are:
        ///     image/gif
        ///     image/jpg
        ///     image/pjpeg
        ///     image/tif
        ///     image/tiff
        ///     text/html
        ///     application/msword
        ///     application/octet-stream
        ///     application/vnd.openxmlformats-officedocument.wordprocessingml.document
        /// </summary>
        /// <returns>Returns true if the conversion succeeds. Otherwise, false and the message property contains the log.</returns>
        public bool Convert()
        {
            bool isConverted = false;
            // Flag to checked if we need to convert the document to PDF.
            bool needsConversion = false;
            IConverter converter = null;
            Message = string.Empty;

            try
            {
                // Ensure that the file exists.
                if (File.Exists(InputFile))
                {
                    switch (MimeType)
                    {
                        case Constant.MimeType.IMAGE_GIF:
                        case Constant.MimeType.IMAGE_JPG:
                        case Constant.MimeType.IMAGE_PJPEG:
                            {
                                needsConversion = true;
                                converter = new ImageConverter();
                                break;
                            }
                        case Constant.MimeType.IMAGE_TIF:
                        case Constant.MimeType.IMAGE_TIFF:
                            {
                                needsConversion = true;
                                converter = new TIFFConverter();
                                break;
                            }
                        case Constant.MimeType.TEXT_HTML:
                            {
                                needsConversion = true;
                                converter = new HTMLConverter();
                                break;
                            }
                        case Constant.MimeType.APPLICATION_MSWORD:
                        case Constant.MimeType.APPLICATION_OCTET_STREAM:
                        case Constant.MimeType.APPLICATION_OPENXMLFORMATS_OFFICEDOCUMENT:
                            {
                                needsConversion = true;
                                converter = new WordConverter();  
                                break;
                            }
                        case Constant.MimeType.APPLICATION_PDF:
                            {
                                // PDFs are not converted.
                                needsConversion = false;
                                OutputFile = InputFile;
                                break;
                            }
                        default: // Unsupported mimetype which cannot be converted.
                            {
                                // Log unsupported mimetype.
                                needsConversion = false;
                                Message = "File '"
                                            + InputFile + "' cannot be converted to PDF because the file type is not supported (see file extension).";
                                break;
                            }
                    }

                    if (needsConversion)
                    {
                        isConverted = converter.ConvertToPDF(InputFile, OutputFile);
                        if (isConverted)
                        {
                            Message = "Successful PDF conversion for file "
                                        + InputFile
                                        + ".";
                        }
                        else
                        {
                            // Log failed conversion.
                            Message = "Failed PDF conversion for file "
                                        + InputFile
                                        + ".";
                        }
                    }
                }
                else
                {
                    // Log failed copy to temporary folder.
                    Message = "Input file '" + InputFile + "' is missing, unable to convert.";
                }
            }
            catch (Exception ex)
            {
                Message = "An exception occurred while converting " + InputFile + " to PDF."
                            + Environment.NewLine
                            + "Exception message: "
                            + Environment.NewLine
                            + ex.Message;
                isConverted = false;
            }
            finally
            { 
            
            }

            return isConverted;
        }
    }
}
