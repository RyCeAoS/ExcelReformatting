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
    public interface IReadService
    {
        public  Task ReadExcelsheet(IFormFile file, List<object> Output);

    }
}
