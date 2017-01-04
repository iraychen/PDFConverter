using HtmlAgilityPack;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace HealthShare.PDFConverter
{
    public class CustomImageProvider : AbstractImageProvider
    {
        string imagePath = "Images/";

        public CustomImageProvider()
            : base()
        { }

        public CustomImageProvider(string path)
            : base()
        {
            imagePath = path;
        }


        public override string GetImageRootPath()
        {
            return imagePath;
        }

        public override void Reset()
        { }

        public override iTextSharp.text.Image Retrieve(string src)
        {
            return (iTextSharp.text.Image)base.Retrieve(src);
        }

        public override void Store(string src, iTextSharp.text.Image img)
        {
            base.Store(src, img);
        }

    }

    public class HTMLConverter : IConverter
    {
        public bool ConvertToPDF(string inputFile, string outputFile)
        {
            bool converted = false;

            try
            {
                HtmlDocument doc = new HtmlDocument();
                doc.OptionFixNestedTags = true;
                doc.OptionWriteEmptyNodes = true;
                doc.OptionAutoCloseOnEnd = true;

                doc.Load(inputFile);

                string rootInner = doc.DocumentNode.InnerHtml;

                if (!rootInner.Contains("<html"))
                {
                    doc.DocumentNode.InnerHtml = "<!DOCTYPE html>"
                                                    + "\r\n<html lang=\"en\" xmlns=\"http://www.w3.org/1999/xhtml\">"
                                                    + "\r\n\t<head>"
                                                    + "\r\n\t\t<title>HTML to PDF</title>"
                                                    + "\r\n\t</head>"
                                                    + "\r\n\t<body>\r\n"
                                                    + rootInner
                                                    + "\r\n\t</body>"
                                                    + "\r\n</html>";
                }

                rootInner = doc.DocumentNode.InnerHtml;

                // Remove <meta> tag.
                string metaTag = @"<\s*(meta)(\s[^>]*)?>\s*";
                while (Regex.IsMatch(rootInner, metaTag))
                {
                    string metaMatch = Regex.Match(rootInner, metaTag).Value;
                    rootInner = rootInner.Replace(metaMatch, string.Empty);
                }
                rootInner = rootInner.Replace("</meta>", string.Empty);

                // Remove <form> tag.
                string formTag = @"<\s*(form)(\s[^>]*)?>\s*";
                while (Regex.IsMatch(rootInner, formTag))
                {
                    string formMatch = Regex.Match(rootInner, formTag).Value;
                    rootInner = rootInner.Replace(formMatch, string.Empty);
                }
                rootInner = rootInner.Replace("</form>", string.Empty);

                // Close br tag.
                string brTag = @"<\s*(br)(\s[^>]*)?>";
                MatchCollection brMatches = Regex.Matches(rootInner, brTag);
                if (brMatches != null)
                {
                    foreach (Match match in brMatches)
                    {
                        rootInner = rootInner.Replace(match.Value, "<br />");
                    }
                }

                // Replace <font> tag with div.
                string fontTag = @"<\s*(font)(\s[^>]*)?>";
                while (Regex.IsMatch(rootInner, fontTag))
                {
                    string fontMatch = Regex.Match(rootInner, fontTag).Value;
                    string toSpan = fontMatch.Replace("font", "div");
                    rootInner = rootInner.Replace(fontMatch, toSpan);
                }
                rootInner = rootInner.Replace("</font>", "</div>");

                doc.DocumentNode.InnerHtml = rootInner;

                // Table elements
                var tableNodes = doc.DocumentNode.SelectNodes("//table");
                if (tableNodes != null)
                {
                    foreach (HtmlNode node in tableNodes)
                    {
                        bool isValidTable = false;
                        if (node.HasChildNodes)
                        {
                            if ((node.SelectSingleNode("thead/tr/th") != null)
                                || (node.SelectSingleNode("tbody/tr/td") != null)
                                || (node.SelectSingleNode("tr/td") != null))
                            {
                                isValidTable = true;
                            }
                        }

                        // Remove invalid table (no tr, td tags).
                        if (!isValidTable)
                        {
                            HtmlNode parent = node.ParentNode;
                            parent.InnerHtml = " ";
                        }

                        bool hasStyle = node.Attributes.Contains("style");
                        if (!hasStyle)
                        {
                            node.Attributes.Add("style", string.Empty);
                        }

                        StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                        if (!string.IsNullOrEmpty(styleValue.ToString()))
                        {
                            if (!styleValue.ToString().EndsWith(";"))
                            {
                                styleValue.Append(";");
                            }
                        }

                        if (node.Attributes.Contains("cellspacing"))
                        {
                            if (!styleValue.ToString().Contains("border-collapse"))
                            {
                                styleValue.Append("border-collapse: collapse; ");
                            }
                            node.Attributes.Remove("cellspacing");
                        }

                        if (node.Attributes.Contains("cellpadding"))
                        {
                            if (!styleValue.ToString().Contains("cellpadding"))
                            {
                                styleValue.Append("padding: " + node.Attributes["cellpadding"].Value + "px; ");
                            }
                            node.Attributes.Remove("cellpadding");
                        }

                        if (node.Attributes.Contains("border"))
                        {
                            if (!styleValue.ToString().Contains("border"))
                            {
                                styleValue.Append("border: " + node.Attributes["border"].Value + "; ");
                            }
                            node.Attributes.Remove("border");
                        }

                        if (node.Attributes.Contains("width"))
                        {
                            string width = node.Attributes["width"].Value;
                            if (node.Attributes["width"].Value.EndsWith("%"))
                            {
                                width = node.Attributes["width"].Value;
                            }
                            else
                            {
                                if (node.Attributes["width"].Value.EndsWith("px"))
                                {
                                    width = node.Attributes["width"].Value;
                                }
                                else
                                {
                                    width = node.Attributes["width"].Value + "px";
                                }
                            }

                            styleValue.Append("width: " + width + "; ");
                            node.Attributes.Remove("width");
                        }

                        if (node.Attributes.Contains("height"))
                        {
                            string height = node.Attributes["height"].Value.EndsWith("px") ? node.Attributes["height"].Value : node.Attributes["height"].Value + "px; ";
                            styleValue.Append("height: " + height);
                            node.Attributes.Remove("height");
                        }

                        if (node.Attributes.Contains("align"))
                        {
                            styleValue.Append("text-align: " + node.Attributes["align"].Value + "; ");
                            node.Attributes.Remove("align");
                        }

                        node.Attributes["style"].Value = styleValue.ToString();
                    }
                }

                // Remove div from /div/img path.
                var imgNodes = doc.DocumentNode.SelectNodes("//div/img");
                while (imgNodes != null)
                {
                    foreach (HtmlNode node in imgNodes)
                    {
                        node.InnerHtml = " ";
                        HtmlNode td = node.ParentNode.ParentNode;
                        td.RemoveChild(node.ParentNode, true);
                    }
                    imgNodes = doc.DocumentNode.SelectNodes("//div/img");
                }

                // Remove div with class="Top_Hidden".
                var divHiddenNodes = doc.DocumentNode.SelectNodes("//div[@class='Top_Hidden']");
                while (divHiddenNodes != null)
                {
                    foreach (HtmlNode node in divHiddenNodes)
                    {
                        HtmlNode tatay = node.ParentNode;
                        tatay.RemoveChild(node, false);
                    }
                    divHiddenNodes = doc.DocumentNode.SelectNodes("//div[@class='Top_Hidden']");
                }

                // Remove children for div with class="blank".
                var divBlankNodes = doc.DocumentNode.SelectNodes("//div[@class='blank']");
                if (divBlankNodes != null)
                {
                    foreach (HtmlNode node in divBlankNodes)
                    {
                        node.RemoveAllChildren();
                    }
                }

                // Close img tag from /td/img path.
                var tdImgNodes = doc.DocumentNode.SelectNodes("//td/img");
                if (tdImgNodes != null)
                {
                    foreach (HtmlNode node in tdImgNodes)
                    {
                        node.InnerHtml = " ";
                    }
                }

                // Add style to div tag.
                var divNodes = doc.DocumentNode.SelectNodes("//div");
                if (divNodes != null)
                {
                    foreach (HtmlNode node in divNodes)
                    {
                        if (node.HasAttributes)
                        {
                            bool hasStyle = node.Attributes.Contains("style");
                            if (!hasStyle)
                            {
                                node.Attributes.Add("style", string.Empty);
                            }

                            StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                            if (!string.IsNullOrEmpty(styleValue.ToString()) && !styleValue.ToString().EndsWith(";"))
                            {
                                styleValue.Append(";");
                            }

                            if (node.Attributes.Contains("face"))
                            {
                                string fontFamily = node.Attributes["face"].Value;
                                styleValue.Append("font-family: " + fontFamily.ToLower() + ";");
                                node.Attributes.Remove("face");
                            }

                            if (node.Attributes.Contains("size"))
                            {
                                string fontSize = node.Attributes["size"].Value;
                                string size = "9pt";
                                switch (fontSize)
                                {
                                    case "1":
                                        {
                                            size = "7pt";
                                            break;
                                        }
                                    case "2":
                                        {
                                            size = "9pt";
                                            break;
                                        }
                                    case "3":
                                        {
                                            size = "10pt";
                                            break;
                                        }
                                    case "4":
                                        {
                                            size = "12pt";
                                            break;
                                        }
                                    case "5":
                                        {
                                            size = "16pt";
                                            break;
                                        }
                                    case "6":
                                        {
                                            size = "20pt";
                                            break;
                                        }
                                    case "7":
                                        {
                                            size = "30pt";
                                            break;
                                        }
                                    default:
                                        break;
                                }

                                styleValue.Append("font-size: " + size.ToLower() + ";");
                                node.Attributes.Remove("size");
                            }

                            node.Attributes["style"].Value = styleValue.ToString();
                        }
                    }
                }

                // Add td style.
                var tdNodes = doc.DocumentNode.SelectNodes("//td");
                if (tdNodes != null)
                {
                    foreach (HtmlNode node in tdNodes)
                    {
                        bool hasStyle = node.Attributes.Contains("style");
                        if (!hasStyle)
                        {
                            node.Attributes.Add("style", string.Empty);
                        }

                        StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                        if (!string.IsNullOrEmpty(styleValue.ToString()))
                        {
                            if (!styleValue.ToString().EndsWith(";"))
                            {
                                styleValue.Append("; ");
                            }
                        }

                        if (node.Attributes.Contains("align"))
                        {
                            styleValue.Append("text-align: " + node.Attributes["align"].Value + "; ");
                            node.Attributes.Remove("align");
                        }
                        else
                        {
                            styleValue.Append("text-align: left;");
                        }

                        if (node.Attributes.Contains("valign"))
                        {
                            styleValue.Append("vertical-align: " + node.Attributes["valign"].Value + "; ");
                            node.Attributes.Remove("valign");
                        }

                        if (node.Attributes.Contains("width"))
                        {
                            string width = node.Attributes["width"].Value;
                            if (node.Attributes["width"].Value.EndsWith("%"))
                            {
                                width = node.Attributes["width"].Value;
                            }
                            else
                            {
                                if (node.Attributes["width"].Value.EndsWith("px"))
                                {
                                    width = node.Attributes["width"].Value;
                                }
                                else
                                {
                                    width = node.Attributes["width"].Value + "px";
                                }
                            }

                            styleValue.Append("width: " + width + "; ");
                            node.Attributes.Remove("width");
                        }

                        if (!string.IsNullOrEmpty(styleValue.ToString()))
                        {
                            node.Attributes["style"].Value = styleValue.ToString();
                        }
                    }
                }

                // Add style to p tag.
                var pNodes = doc.DocumentNode.SelectNodes("//p");
                if (pNodes != null)
                {
                    foreach (HtmlNode node in pNodes)
                    {
                        if (node.HasAttributes)
                        {
                            bool hasStyle = node.Attributes.Contains("style");
                            if (!hasStyle)
                            {
                                node.Attributes.Add("style", string.Empty);
                            }

                            StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                            if (!string.IsNullOrEmpty(styleValue.ToString()) && !styleValue.ToString().EndsWith(";"))
                            {
                                styleValue.Append(";");
                            }

                            if (node.Attributes.Contains("align"))
                            {
                                string value = node.Attributes["align"].Value;
                                styleValue.Append("text-align: " + value.ToLower() + ";");
                                node.Attributes.Remove("align");
                            }

                            node.Attributes["style"].Value = styleValue.ToString();
                        }
                    }
                }

                // Remove u tag from //u/div path but put underline in div tag.
                var uDivNodes = doc.DocumentNode.SelectNodes("//u/div");
                while (uDivNodes != null)
                {
                    foreach (HtmlNode node in uDivNodes)
                    {
                        bool hasStyle = node.Attributes.Contains("style");
                        if (!hasStyle)
                        {
                            node.Attributes.Add("style", string.Empty);
                        }

                        StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                        if (!string.IsNullOrEmpty(styleValue.ToString()) && !styleValue.ToString().EndsWith(";"))
                        {
                            styleValue.Append(";");
                        }

                        styleValue.Append("text-decoration: underline;");

                        node.Attributes["style"].Value = styleValue.ToString();

                        HtmlNode lolo = node.ParentNode.ParentNode;
                        lolo.RemoveChild(node.ParentNode, true);
                    }
                    uDivNodes = doc.DocumentNode.SelectNodes("//u/div");
                }

                // Remove strong tag from //strong/div path but put bold in div tag.
                var strongDivNodes = doc.DocumentNode.SelectNodes("//strong/div");
                while (strongDivNodes != null)
                {
                    foreach (HtmlNode node in strongDivNodes)
                    {
                        bool hasStyle = node.Attributes.Contains("style");
                        if (!hasStyle)
                        {
                            node.Attributes.Add("style", string.Empty);
                        }

                        StringBuilder styleValue = new StringBuilder(node.Attributes["style"].Value.Trim());
                        if (!string.IsNullOrEmpty(styleValue.ToString()) && !styleValue.ToString().EndsWith(";"))
                        {
                            styleValue.Append(";");
                        }

                        styleValue.Append("font-weight: bold;");

                        node.Attributes["style"].Value = styleValue.ToString();

                        HtmlNode lolo = node.ParentNode.ParentNode;
                        lolo.RemoveChild(node.ParentNode, true);
                    }
                    strongDivNodes = doc.DocumentNode.SelectNodes("//strong/div");
                }

                // Replace p tag with div from //p/div path.
                var pDivNodes = doc.DocumentNode.SelectNodes("//p/div");
                while (pDivNodes != null)
                {
                    foreach (HtmlNode node in pDivNodes)
                    {
                        node.ParentNode.Name = "div";
                    }
                    pDivNodes = doc.DocumentNode.SelectNodes("//p/div");
                }


                // Remove div tag from //ol/div path.
                var olDivNodes = doc.DocumentNode.SelectNodes("//ol/div");
                while (olDivNodes != null)
                {
                    foreach (HtmlNode node in olDivNodes)
                    {
                        HtmlNode tatay = node.ParentNode;

                        bool hasStyle = node.Attributes.Contains("style");
                        if (hasStyle)
                        {
                            tatay.Attributes.Add(node.Attributes["style"]);
                        }

                        tatay.RemoveChild(node, true);
                    }
                    olDivNodes = doc.DocumentNode.SelectNodes("//ol/div");
                }

                // Remove div tag from //ul/div path.
                var ulDivNodes = doc.DocumentNode.SelectNodes("//ul/div");
                while (ulDivNodes != null)
                {
                    foreach (HtmlNode node in ulDivNodes)
                    {
                        HtmlNode tatay = node.ParentNode;

                        bool hasStyle = node.Attributes.Contains("style");
                        if (hasStyle)
                        {
                            tatay.Attributes.Add(node.Attributes["style"]);
                        }

                        tatay.RemoveChild(node, true);
                    }
                    ulDivNodes = doc.DocumentNode.SelectNodes("//ul/div");
                }

                // Remove div tag from //li/div path.
                var liDivNodes = doc.DocumentNode.SelectNodes("//li/div");
                while (liDivNodes != null)
                {
                    foreach (HtmlNode node in liDivNodes)
                    {
                        HtmlNode tatay = node.ParentNode;

                        bool hasStyle = node.Attributes.Contains("style");
                        if (hasStyle)
                        {
                            tatay.Attributes.Add(node.Attributes["style"]);
                        }

                        tatay.RemoveChild(node, true);
                    }
                    liDivNodes = doc.DocumentNode.SelectNodes("//li/div");
                }

                // Save the modified html to a new name.
                string formattedHTMLFile = outputFile.Replace(".pdf", string.Empty) + "_TEMP.html";
                doc.Save(formattedHTMLFile);

                HtmlDocument docCPM = new HtmlDocument();
                docCPM.OptionFixNestedTags = true;
                docCPM.OptionWriteEmptyNodes = true;
                docCPM.OptionAutoCloseOnEnd = true;

                string newHTML = outputFile.Replace(".pdf", string.Empty) + "_PDF.html";
                docCPM.Load(formattedHTMLFile);
                docCPM.Save(newHTML);

                FontFactory.RegisterDirectories();
                Document document = new Document();
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outputFile, FileMode.Create));
                try
                {
                    document.Open();

                    HtmlPipelineContext htmlContext = new HtmlPipelineContext(null);
                    htmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory());
                    CustomImageProvider imageProvider = new CustomImageProvider();
                    htmlContext.SetImageProvider(imageProvider);

                    ICSSResolver cssResolver = XMLWorkerHelper.GetInstance().GetDefaultCssResolver(true);

                    IPipeline pipeline =
                        new CssResolverPipeline(cssResolver,
                            new HtmlPipeline(htmlContext,
                                new PdfWriterPipeline(document, writer)));

                    XMLWorker worker = new XMLWorker(pipeline, true);
                    XMLParser p = new XMLParser(worker);
                    p.Parse(new StreamReader(newHTML));
                    //XMLWorkerHelper.GetInstance().ParseXHtml(writer, document, new StreamReader(newHTMLFile));

                    converted = true;
                }
                catch
                {
                    converted = false;
                    throw;
                }
                finally
                {
                    if (document.IsOpen())
                    {
                        document.Close();
                    }

                    if (writer != null)
                    { 
                        writer.Close();
                    }

                    File.Delete(formattedHTMLFile);
                    File.Delete(newHTML);
                }
            }
            catch
            {
                converted = false;
                throw;
            }
            finally
            {
            }

            return converted;
        }
    }

}
