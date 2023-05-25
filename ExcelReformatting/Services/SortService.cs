using OfficeOpenXml;
using ExcelReformatting.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
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
    public class SortService : ISortService
    {

        List<Client> noalphas = new List<Client>();
        List<Client> alphas = new List<Client>();

        public void SortFile(List<object> clients)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            sorting(clients);


        }


        private void sorting(List<object> clients)
        {

            List<Client> Clients = clients.Cast<Client>().ToList();

            foreach (Client client in Clients)
            {
                bool hasAlpha = client.c_n.Any(x => char.IsLetter(x));
                if (hasAlpha)
                {
                    trimming(client);
                }
                if (!hasAlpha)
                {
                    noalphas.Add(client);
                }
            }


            alphas = Clients;
            quickSort(noalphas, 0, (noalphas.Count) - 1);
            quickSort(alphas, 0, (alphas.Count) - 1);
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




        private void trimming(Client client)
        {
            string newCaseNumber = "";
            if (client.c_n.Length < 10 && client.c_n != null)
            {
                int number_of_zeros = 10 - client.c_n.Length;
                for (int i = 0; i < number_of_zeros; i++)
                {
                    newCaseNumber = client.c_n.Insert(0, "0");
                }

                client.c_n = newCaseNumber;
            }

            else if (client.c_n.Length > 10 && client.c_n != null)
            {
                int num_to_cut = client.c_n.Length - 10;
                newCaseNumber = client.c_n.Remove(0, num_to_cut);
                client.c_n = newCaseNumber;
            }
        }

        private void loadfiles(List<ExcelPackage> files, ExcelPackage file, List<Client> clients, String wsName)
        {
            var ws = file.Workbook.Worksheets.Add(Name: wsName);
            var range = ws.Cells["A1"].LoadFromCollection(clients, PrintHeaders:true );
            range.AutoFitColumns();
            ws.Row(row: 1).Style.Font.Bold = true;
            files.Add(file);
        }






        private List<ExcelPackage> GetPackages()
        {
            List<ExcelPackage> packages = new List<ExcelPackage>();
            ExcelPackage file1 = new ExcelPackage();
            ExcelPackage file2 = new ExcelPackage();

            loadfiles(packages, file1, alphas, "MainReport");
            loadfiles(packages, file2, noalphas, "NoAlphas");


            var ws = packages[0].Workbook.Worksheets[PositionID: 0];
            ws.InsertColumn(4, 1);
            ws.Cells["D1"].Value = "New Case number";
            ws.Cells["D1"].Style.Font.Bold = true;
            


            return packages;
        }



        private List<string> GetFileNames()
        {
            List<string> names = new List<string>();
            names.Add("MainReport.xlsx");
            names.Add("NoAlphas.xlsx");
            


            return names;
        }

      


   


        public async Task<FileContentResult> GetFiles()
        {
            using (var ms = new MemoryStream())
            {
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    byte[] bytes1 = GetPackages()[0].GetAsByteArray();
                    byte[] bytes2 = GetPackages()[1].GetAsByteArray();
                    

                    var zipEntry = archive.CreateEntry(GetFileNames()[0], System.IO.Compression.CompressionLevel.Fastest);
                    using (var zipStream = zipEntry.Open())
                    {
                        await zipStream.WriteAsync(bytes1, 0, bytes1.Length);
                    }

                    var zipEntry2 = archive.CreateEntry(GetFileNames()[1], System.IO.Compression.CompressionLevel.Fastest);
                    using (var zipStream = zipEntry2.Open())
                    {
                        await zipStream.WriteAsync(bytes2, 0, bytes2.Length);
                    }


                  
                }

                return new FileContentResult(ms.ToArray(), "application/zip");
            }
        }


        public async Task<FileContentResult> GetFile()
        {
            return await GetFiles();
        }

    } 
}
