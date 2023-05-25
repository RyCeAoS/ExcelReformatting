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

namespace ExcelReformatting.Interfaces
{
    public interface ICaseNumberDoc
    {
        public List<String> CaseNumbersToList(List<ClientDoc> Clients);

        public MemoryStream WordDoc(List<String> case_numbers);

        public Task<FileContentResult> Case_numbers_doc(MemoryStream ms);
    }
}
