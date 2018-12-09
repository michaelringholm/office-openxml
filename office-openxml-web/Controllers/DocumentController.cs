
using System;
using System.Collections.Generic;
using com.opusmagus.azure.graph;
using com.opusmagus.office.openxml;
using com.opusmagus.cloud.blobs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace office_openxml_web.Pages
{
    [Route("api/doc/")]
    [ApiController]
    public class DocumentController : ControllerBase
    {
        private IOpenDocument openDocument;
        private IDocumentProvider docProvider;
        private string blobContainerName;
        private string blobContainerConnString;

        public DocumentController(IOpenDocument openDocument, IDocumentProvider docProvider, IConfiguration configuration)
        {
            this.openDocument = openDocument;
            this.docProvider = docProvider;
            blobContainerName = configuration.GetSection("blobContainerName").Value; //"sample-blob-container";
            blobContainerConnString = configuration.GetSection("blobContainerConnString").Value;
        }

        /*public DocumentController(IOpenDocument openDocument, IDocumentProvider docProvider, string blobContainerName, string blobContainerConnString)
        {
            this.openDocument = openDocument;
            this.docProvider = docProvider;
            this.blobContainerName = blobContainerName;
            this.blobContainerConnString = blobContainerConnString;
        }*/       

        [Route("preview/{pdfGuid}"), AcceptVerbs("Get")]
        public ActionResult preview(string pdfGuid)
        {
            //string serverPath = Server.MapPath(filepath);
            var path = $"./local/{pdfGuid} - REPLACED.pdf";
            var bytes = System.IO.File.ReadAllBytes(path);
            return File(bytes, "application/pdf");
        }

        //[HttpPost]
        [Route("modify"), AcceptVerbs("POST")]
        public string modify([FromBody] DocumentHeader documentHeader)
        {
            IBlobService blobService = new AzureBlobService();
            var blobContents = blobService.getBlobContents(blobContainerConnString, blobContainerName, documentHeader.BlobItemName);
            var tempGUID = Guid.NewGuid().ToString();
            var tempDocPath = $"./local/{tempGUID}";
            System.IO.Directory.CreateDirectory(tempDocPath);
            System.IO.File.WriteAllBytes($"{tempDocPath}.docx", blobContents);

            var bookmarkReplacements = new Dictionary<string, string>();
            bookmarkReplacements.Add("Commentor_adresse", "Andevej 14");
            bookmarkReplacements.Add("Commentor_navn_header", "Anders And");
            bookmarkReplacements.Add("Commentor_navn_body", "Anders And");
            bookmarkReplacements.Add("Commentor_registreringsnummer", "1234 1234512345");
            bookmarkReplacements.Add("Commentor_dato", DateTime.Now.ToString("dd.MM.yyyy"));
            bookmarkReplacements.Add("Commentor_bodytext", documentHeader.BodyText);
            openDocument.ReplaceBookmarks($"{tempDocPath}.docx", $"{tempDocPath} - REPLACED.docx", bookmarkReplacements);

            var htmlSimpleInputBytes = System.IO.File.ReadAllBytes($"{tempDocPath} - REPLACED.docx");
            var serviceUser = docProvider.GetUser("pdf@commentor.dk");
            byte[] pdfHTMLSimpleDocBytes = docProvider.ConvertDocumentToPDF(htmlSimpleInputBytes, $"Temp/{Guid.NewGuid()}.docx", serviceUser.Id);
            System.IO.File.WriteAllBytes($"{tempDocPath} - REPLACED.pdf", pdfHTMLSimpleDocBytes);

            return tempGUID;
        }

        [Route("templates"), AcceptVerbs("Get")]
        public List<String> getTemplates()
        {
            IBlobService blobService = new AzureBlobService();            
            var blobItems = blobService.getBlobItems(blobContainerConnString, blobContainerName, 10);
            var blobNames = new List<String>();
            foreach(var blobItem in blobItems)
                blobNames.Add(((CloudBlockBlob)blobItem).Name);

            return blobNames;
        }        
    }
}