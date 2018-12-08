using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace com.opusmagus.office.openxml
{
    public class OpenDocument
    {
        public static void ReplaceProperties(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open("./resources/source/properties.docx", true))
            {
                var body = doc.MainDocumentPart.Document.Body;                

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    Console.WriteLine(paragraph);
                    /*if (paragraph.Text.Contains("##sagsid##"))
                    {
                        text.Text = text.Text.Replace("##sagsid##", "Sagsid: 199");
                    }*/
                }                
                
                foreach (var text in body.Descendants<Text>())
                {
                    if (text.Text.Contains("##sagsid##"))
                    {
                        text.Text = text.Text.Replace("##sagsid##", "Sagsid: 199");
                    }
                }
              
                foreach (var customProp in doc.CustomFilePropertiesPart.Properties.Descendants<CustomDocumentProperty>())
                {
                    Console.WriteLine($"Name={customProp.Name} Value={customProp.InnerText}");
                    //customProp.SetAttribute(customProp);
                    //customProp = 173;
                }

                doc.Save();
            }
        }
    
        public static void ReplaceBookmarks(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks)
        {
            bool isEditable = true;

            // Don't use open unless you want to change that document too, use CreateFromTemplate instead
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(sourceDocPath, isEditable))
            {
                var body = doc.MainDocumentPart.Document.Body;                


                foreach (var bookmarkStart in body.Descendants<BookmarkStart>())
                {
                    Console.WriteLine($"bookmarkName={bookmarkStart.Name}");
                    Console.WriteLine($"bookmarkValue={bookmarkStart.InnerText}");
                    if (bookmarkStart.Name.InnerText.Equals("sagsnrtitel")) {
                        //var text = bookmarkStart.Descendants<Text>();
                        //text.
                        var sibling = bookmarkStart.NextSibling();
                        //(new Text("1888"));
                    }
                    if (bookmarkStart.Name.InnerText.Equals("sagsnr")) {
                        //var text = bookmarkStart.Descendants<Text>();
                        //text.
                        var sibling = bookmarkStart.NextSibling();
                        //sibling.AppendChild()
                        //(new Text("1888"));
                    }
                    if (bookmarkStart.Name.InnerText.Equals("faxnr")) {
                        //var text = bookmarkStart.Descendants<Text>();
                        //text.
                        var bookmarkEnd = bookmarkStart.NextSibling<BookmarkEnd>();
                        var bookmarkRun = new Run();
                        bookmarkStart.Parent.InsertAfter(bookmarkRun, bookmarkStart);
                        bookmarkRun.AppendChild(new Text("26 83 68 98"));
                        //var bookmarkRun = bookmarkStart.NextSibling<Run>();
                        //(new Text("1888"));
                    }    
                    if (bookmarkStart.Name.InnerText.Equals("mobilnr")) {
                        //var text = bookmarkStart.Descendants<Text>();
                        //text.
                        var bookmarkEnd = bookmarkStart.NextSibling<BookmarkEnd>();
                        var bookmarkRun = bookmarkStart.NextSibling<Run>();
                        var xmlElement = bookmarkRun.GetFirstChild<OpenXmlElement>();
                        bookmarkRun.AppendChild(new Text("26 83 69 97"));
                        bookmarkRun.RemoveChild(xmlElement);
                        //bookmarkRun.InnerText = "test";
                        //(new Text("1888"));
                    }     
                    if (bookmarkStart.Name.InnerText.StartsWith("Commentor_")) {
                        //var text = bookmarkStart.Descendants<Text>();
                        //text.
                        var bookmarkEnd = bookmarkStart.NextSibling<BookmarkEnd>();
                        var bookmarkRun = bookmarkStart.NextSibling<Run>();
                        var xmlElement = bookmarkRun.GetFirstChild<OpenXmlElement>();
                        //bookmarkRun.AppendChild(new Text("26 83 69 97"));
                        //bookmarkRun.RemoveChild(xmlElement);
                        //bookmarkRun.InnerText = "test";
                        //(new Text("1888"));
                    }                                
                }                

                doc.SaveAs(targetDocPath);
            }
        }
    }
}