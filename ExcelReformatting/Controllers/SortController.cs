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
    public class SortController : Controller
    {
        List<Client> output = new List<Client>(); //list of neighbors that will be imported from the excel sheet
        List<Client> noalphas = new List<Client>();
        List<Client> alphas = new List<Client>();
        public IActionResult Index()
        {
            return View();
        }



        public async Task<FileContentResult> SortedFile(IFormFile file)
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
                        /*  a  */           c.wfid = ws.Cells[row, col].Value.ToString();
                        /*  b  */           c.dte = ws.Cells[row, col + 1].Value.ToString();
                        /*  c  */           c.c_n = ws.Cells[row, col + 2].Value.ToString();
                        /*  d  */           c.a_i_d = ws.Cells[row, col + 3].Value.ToString();
                        /*  e  */           c.f_n = ws.Cells[row, col + 4].Value.ToString();
                        /*  f  */           c.ph_num = ws.Cells[row, col + 5].Value.ToString();
                        /*  g  */           c.inM = ws.Cells[row, col + 6].Value.ToString();
                        /*  h  */           c.cntr = ((ws.Cells[row, col + 7].Value) == null) ? null : ws.Cells[row, col + 7].Value.ToString();
                        /*  i  */           c.comp_tpe = ws.Cells[row, col + 8].Value.ToString();
                        /*  j  */           c.comp_desc = ((ws.Cells[row, col + 9].Value) == null) ? null : ws.Cells[row, col + 9].Value.ToString();
                        /*  k  */           c.comp_res = ((ws.Cells[row, col + 10].Value) == null) ? null : ws.Cells[row, col + 10].Value.ToString();
                        /*  l  */           c.com = ((ws.Cells[row, col + 11].Value) == null) ? null : ws.Cells[row, col + 11].Value.ToString();
                        /*  m  */           c.Temp = ((ws.Cells[row, col + 12].Value) == null) ? null : ws.Cells[row, col + 12].Value.ToString();
                        /*  n  */           c.Wfc = ((ws.Cells[row, col + 13].Value) == null) ? null : ws.Cells[row, col + 13].Value.ToString();

                        output.Add(c);
                        row++;

                    }
                }
            }

           

          
          foreach(Client client in output)
            {
                char lastLetter = client.c_n[client.c_n.Length - 1];
                if (Char.IsLetter(lastLetter))
                {
                    alphas.Add(client);
                }
                else
                {
                    noalphas.Add(client);
                }
            }

            quickSort(noalphas, 0, (noalphas.Count) - 1);
            quickSort(alphas, 0, (alphas.Count) - 1);
        }

        private async Task<FileContentResult> SaveExcelsheet()
        {
            using (ExcelPackage resultantPackage = new ExcelPackage())
            {
                
                var ws1 = resultantPackage.Workbook.Worksheets.Add(Name: "Alphas");
                var ws2 = resultantPackage.Workbook.Worksheets.Add(Name: "No Alphas");

                var range1 = ws1.Cells["A1"].LoadFromCollection(alphas, PrintHeaders: true);
                var range2 = ws2.Cells["A1"].LoadFromCollection(noalphas, PrintHeaders: true);

                range1.AutoFitColumns();
                range2.AutoFitColumns();

                ws1.Row(row: 1).Style.Font.Bold = true;
                ws2.Row(row: 1).Style.Font.Bold = true;

                FileContentResult result_excel_file = await Task.Run(() => new FileContentResult(resultantPackage.GetAsByteArray(), " application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                return result_excel_file;
            }
        }


        private static void quickSort(List<Client> list, int start, int end)
        {

           if (end <= start) return; //base case

            int pivot = partition(list, start, end);
            quickSort(list, start, pivot - 1);
            quickSort(list, pivot + 1, end);
        }
        private static int partition(List<Client> list, int start, int end)
        {

            Client pivot = list[end];
            int i = start - 1;
            Client temp;

            //while going throught the array, if j is smaller than the pivot value at the end then we increment i and swap it with j
            for (int j = start; j <= end; j++)
            {
                if (String.Compare(list[j].c_n, pivot.c_n) < 0)
                {
                    i++;
                    temp = list[i];
                    list[i] = list[j];
                    list[j] = temp;
                }
            }

            //When reach the end of the initial array we swap the pivot value where i rests
            i++;
            temp = list[i];
            list[i] = list[end];
            list[end] = temp;

            return i;
        }

    }
}
