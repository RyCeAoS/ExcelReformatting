using OfficeOpenXml;
using ExcelReformatting.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using ExcelReformatting.Service;
using ExcelReformatting.Interfaces;

namespace ExcelReformatting.Controllers
{
    public class ServiceController : Controller
    {
        private readonly IReadService _read;
        private readonly IReadService _read2;
        private readonly ISortService _sort;
        private readonly ICaseNumberDoc _case;


        public ServiceController(Startup.ServiceResolver serviceResolver, ISortService sort, ICaseNumberDoc caseNumberDoc) 
        {
            _read = serviceResolver("A");
            _read2 = serviceResolver("B");
            _case = caseNumberDoc;
            _sort = sort;
        }

        public async Task<FileContentResult> getfile(IFormFile file)
        {
            List<object> clients = new List<object>();
            await _read2.ReadExcelsheet(file, clients);
            List<ClientDoc> client_docs = clients.Cast<ClientDoc>().ToList();
            List<String> casenumbers = _case.CaseNumbersToList(client_docs);
            var ms = _case.WordDoc(casenumbers);
            return await _case.Case_numbers_doc(ms);
        }
            


        public async Task<FileContentResult> FormattedExcelFile(IFormFile file)
        {
            List<object> clients = new List<object>();
            await _read.ReadExcelsheet(file, clients);
            _sort.SortFile(clients);
            var savesheet = await _sort.GetFile();
            return savesheet;
        }
    }
}
