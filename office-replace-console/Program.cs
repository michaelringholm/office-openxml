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
            
            OpenDocument.ReplaceProperties("./resources/source/properties.docx", "./resources/target/properties-replaced.docx", null);
            
            var bookmarkReplacements = new Dictionary<string, string>();
            bookmarkReplacements.Add("Commentor_adresse", "Andevej 14");
            bookmarkReplacements.Add("Commentor_navn", "Anders And");
            bookmarkReplacements.Add("Commentor_registreringsnummer", "1234 1234512345");   

            /*var bookmarkReplacements = new Dictionary<string, string>();
            bookmarkReplacements.Add("mobilnr", "26 83 69 97");
            bookmarkReplacements.Add("faxnr", "26 83 68 98");
            OpenDocument.ReplaceBookmarks("./resources/source/bookmarks.docx", "./resources/target/bookmarks-replaced.docx", bookmarkReplacements);*/

            //bookmarkReplacements.Add("Commentor_navn", "26 83 68 98");
            OpenDocument.ReplaceBookmarks("../local/Tekstforslag varslingsbrev december 2018 Version 4.docx", "../local/Tekstforslag varslingsbrev december 2018 Version 4 - REPLACED.docx", bookmarkReplacements);
            
            Console.WriteLine("Office replace demo ended.");
        }
    }
}
