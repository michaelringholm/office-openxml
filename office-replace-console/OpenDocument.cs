using System;
using System.Collections.Generic;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace com.opusmagus.office.openxml
{
    public class OpenDocument : IOpenDocument
    {
        public void ReplaceProperties(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks)
        {
            //using (WordprocessingDocument doc = WordprocessingDocument.Open(sourceDocPath, true))
            var isEditable = true;
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(sourceDocPath, isEditable))
            {
                var body = doc.MainDocumentPart.Document.Body;                

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    Debug.WriteLine(paragraph);
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
                    Debug.WriteLine($"Name={customProp.Name} Value={customProp.InnerText}");
                    //customProp.SetAttribute(customProp);
                    //customProp = 173;
                }

                //doc.Save();
                doc.SaveAs(targetDocPath);
            }
        }
    
        public void ReplaceBookmarks(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks)
        {
            bool isEditable = true;

            // Don't use open unless you want to change that document too, use CreateFromTemplate instead
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(sourceDocPath, isEditable))
            {
                var body = doc.MainDocumentPart.Document.Body;                

                foreach (var bookmarkStart in body.Descendants<BookmarkStart>())
                {
                    Debug.WriteLine($"bookmarkName={bookmarkStart.Name}");
                    Debug.WriteLine($"bookmarkValue={bookmarkStart.InnerText}");
                    var bookmarkName = bookmarkStart.Name.InnerText;
                    if(bookmarks.ContainsKey(bookmarkName))
                        replaceBookmarkText(bookmarkStart, bookmarks.GetValueOrDefault(bookmarkName));
                }                
                doc.SaveAs(targetDocPath);
            }
        }

        private void replaceBookmarkText(BookmarkStart bookmarkStart, string bookmarkValue)
        {
            var bookmarkEnd = bookmarkStart.NextSibling<BookmarkEnd>();
            var bookmarkRun = bookmarkStart.NextSibling<Run>();
            if(bookmarkRun != null)
                bookmarkRun.RemoveAllChildren();
            else {
                bookmarkRun = new Run();
                bookmarkStart.Parent.InsertAfter(bookmarkRun, bookmarkStart);
            }
            bookmarkRun.AppendChild(new Text(bookmarkValue));            
        }
    }
}