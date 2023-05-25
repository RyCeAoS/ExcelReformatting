using OfficeOpenXml;
using ExcelReformatting.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using ExcelReformatting.Models;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using ExcelReformatting.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace ExcelReformatting.Service
{
    public class CaseNumberDoc : ICaseNumberDoc
    {


        public List<String> CaseNumbersToList(List<ClientDoc> Clients)
        {
            List<String> caseNumbers = new List<String>();
            foreach (ClientDoc client in Clients)
            {
                String caseNumber = client.c_n_N;
                caseNumbers.Add(caseNumber);
            }

            return caseNumbers;
        }

        public async Task<FileContentResult> Case_numbers_doc(MemoryStream ms)
        {
            return await Task.Run(() => new FileContentResult(ms.ToArray(), "application/msword"));
        }

        public MemoryStream WordDoc(List<string> case_numbers)
        {
            using (var mem = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    doc.AddMainDocumentPart().Document = new Document();
                    var body = doc.MainDocumentPart.Document.AppendChild(new Body());
                    var paragraph = body.AppendChild(new Paragraph());
                    var run = paragraph.AppendChild(new Run());


                    foreach (string case_number in case_numbers)
                    {
                        run.AppendChild(new Text(case_number + ";"));
                    }
                }

                return mem;
            }


        }
    }
}