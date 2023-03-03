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


namespace ExcelReformatting.Controllers
{
    public class ExcelController : Controller
    {
        List<Client> output = new List<Client>(); //list of neighbors that will be imported from the excel sheet
        public IActionResult Index()
        {
            return View();
        }



        public async Task<FileContentResult> FormattedFile(IFormFile file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            await ReadExcelsheet(file);
            var Saved_Excel_Sheet = await SaveExcelsheet();
            return Saved_Excel_Sheet;
            
        }


        public async Task ReadExcelsheet(IFormFile file)
        {

            // 1). Reads excelsheet
            using (var stream = new MemoryStream())
            {

                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream)) // "using" declaration allows for the package to only be used with in a limited scope
                {


                    var ws = package.Workbook.Worksheets[PositionID: 0]; // the first worksheet within the workbook


                    /* Where the excel sheet will start reading
                     * in this instance it will be from row 2 and column 1 
                     */
                    int row = 2; //will begin on row two
                    int col = 1;
                    int lastrow = ws.Dimension.End.Row;


                    while (row < lastrow)
                    // the file will be read as long as the cell doesn't have a white space or null value 
                    {
                        //columns  
                                        Client c = new Client();
                        /*  a  */       c.wfid = MergedCellvalue(ws, row, col);
                        /*  b  */       c.dte = MergedCellvalue(ws, row, col + 1);
                        /*  c  */       c.c_n = MergedCellvalue(ws, row, col + 2);
                        /*  d  */       c.a_i_d = MergedCellvalue(ws, row, col + 3);
                        /*  e  */       c.f_n = MergedCellvalue(ws, row, col + 4);
                        /*  f  */       c.ph_num = ((ws.Cells[row, col + 5].Value) == null) ? "": ws.Cells[row, col + 5].Value.ToString();
                        /*  g  */       c.inM = ((ws.Cells[row, col + 6].Value) == null) ? "" : ws.Cells[row, col + 6].Value.ToString();
                        /*  h  */       c.cntr = ((ws.Cells[row, col + 7].Value) == null) ? "" : ws.Cells[row, col + 7].Value.ToString();
                        /*  i  */       c.comp_tpe = ((ws.Cells[row, col + 8].Value) == null) ? "" : ws.Cells[row, col + 8].Value.ToString();
                        /*  j  */       c.comp_desc = ((ws.Cells[row, col + 9].Value) == null) ? "" : ws.Cells[row, col + 9].Value.ToString();
                        /*  k  */       c.comp_res = ((ws.Cells[row, col + 10].Value) == null) ? "" : ws.Cells[row, col + 10].Value.ToString();
                        /*  l  */       c.com = ((ws.Cells[row, col + 11].Value) == null) ? "" : ws.Cells[row, col + 11].Value.ToString();
                        /*  m  */       c.Temp = MergedCellvalue(ws, row, col + 12);
                        /*  n  */       c.Wfc = MergedCellvalue(ws,row,col + 13);

                        output.Add(c);
                        row++;

                    }
                }
            }
        }

        private async Task<FileContentResult> SaveExcelsheet()
        {
            using (ExcelPackage resultantPackage = new ExcelPackage())
            {
                var ws = resultantPackage.Workbook.Worksheets.Add(Name: "MainReport");
                var range = ws.Cells["A1"].LoadFromCollection(output, PrintHeaders: true);
                range.AutoFitColumns();
                ws.Row(row: 1).Style.Font.Bold = true;
                FileContentResult result_excel_file =  await Task.Run (() => new FileContentResult(resultantPackage.GetAsByteArray(), " application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                return result_excel_file;
            }
        }



        public string MergedCellvalue(ExcelWorksheet ws, int row, int col)
        {
            var cell = ws.Cells[row, col];
            if (cell.Merge == true)
            {
                var mergedID = ws.MergedCells[row, col]; //returns address of the merged cells
                return ws.Cells[mergedID].First().Value.ToString(); // returns the first value within a sequence 
            }
            else
            {
                if (cell.Value != null) return cell.Value.ToString();
                else return "";
            }
        }
    }
}
