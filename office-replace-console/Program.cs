using System;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace com.opusmagus.office.openxml
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Office replace demo started...");
            
            OpenDocument.ReplaceProperties("./resources/source/bookmarks.docx", "./resources/target/bookmarks-replaced.docx", null);
            
            var bookmarkReplacements = new Dictionary<string, string>();
            bookmarkReplacements.Add("mobilnr", "26 83 69 97");
            bookmarkReplacements.Add("faxnr", "26 83 68 98");
            OpenDocument.ReplaceBookmarks("./resources/source/bookmarks.docx", "./resources/target/bookmarks-replaced.docx", bookmarkReplacements);
            
            Console.WriteLine("Office replace demo ended.");
        }
    }
}
