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
using ExcelReformatting.Interfaces;

namespace ExcelReformatting.Service
{

    public class ReadClientFile : IReadService
    {
        public async Task ReadExcelsheet(IFormFile file, List<object> Output)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //list of neighbors that will be imported from the excel sheet
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
                        ClientDoc c = new ClientDoc();
                        c.c_n_N = MergedCellvalue(ws, row, col + 3);
                       

                        Output.Add(c);
                        row++;

                    }
                }
            }
        }


        private string reformat_date(string date)
        {
            string newDate = "";
            if (date.Length == 20)
                newDate = date.Remove(8, 12);
            else if (date.Length == 21)
                newDate = date.Remove(9, 12);
            else if (date.Length == 22)
                newDate = date.Remove(10, 12);
            else
                return date;

            return newDate;

        }



        private string MergedCellvalue(ExcelWorksheet ws, int row, int col)
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