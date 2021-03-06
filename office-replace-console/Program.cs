﻿using System;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using Microsoft.Extensions.DependencyInjection;

namespace com.opusmagus.office.openxml
{
    class Program
    {
        static void Main(string[] args)
        {
            var startTime = DateTime.Now;
            Console.WriteLine($"Office replace demo started at {startTime}...");
            
            var diContainer = new ServiceCollection();
            diContainer.AddSingleton<IOpenDocument, OpenDocument>();
            var diProvider = diContainer.BuildServiceProvider();
            var openDocument = diProvider.GetService<IOpenDocument>(); 
            
            openDocument.ReplaceProperties("./resources/source/properties.docx", "./resources/target/properties-replaced.docx", null);

            var bookmarkReplacements = new Dictionary<string, string>();
            bookmarkReplacements.Add("Commentor_adresse", "Andevej 14");
            bookmarkReplacements.Add("Commentor_navn_header", "Anders And");
            bookmarkReplacements.Add("Commentor_navn_body", "Anders And");
            bookmarkReplacements.Add("Commentor_registreringsnummer", "1234 1234512345");
            bookmarkReplacements.Add("Commentor_dato", DateTime.Now.ToString("dd.MM.yyyy"));
            openDocument.ReplaceBookmarks("../local/Tekstforslag varslingsbrev december 2018 Version 5.docx", "../local/Tekstforslag varslingsbrev december 2018 Version 5 - REPLACED.docx", bookmarkReplacements);
            
            var endTime = DateTime.Now;
            Console.WriteLine($"Office replace demo ended. Duration was {endTime.Millisecond-startTime.Millisecond}.");
        }
    }
}
