using Crawler.Models;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace Crawler.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {

            var pc = GetPincodes();

            GenerateExcel(pc);
            
            return View();
        }

        private List<Pincode> GetPincodes()
        {
            HtmlDocument document = new HtmlDocument();
            document.Load(@".\Pincode.txt");

            List<List<string>> table = document.DocumentNode.SelectNodes("//table")
            .Descendants("tr")
            .Skip(1)
            .Where(tr => tr.Elements("td").Count() > 1)
            .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
            .ToList();

            List<Pincode> pc = new();
            foreach(var item in table)
            {
                Pincode pincode = new();
                pincode.District = item[1];
                pincode.PinCode = item[3];
                pincode.Latitude = item[5].Replace("&amp;",",").Split(',')[0];
                pincode.Longitude = item[5].Replace("&amp;", ",").Split(',')[1];                
                pc.Add(pincode);
            }

            return pc;
        }

        private void GenerateExcel(List<Pincode> pincode)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {                
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "District";
                excelPackage.Workbook.Properties.Title = "PinCode";
                excelPackage.Workbook.Properties.Subject = "Latitude";
                excelPackage.Workbook.Properties.Comments = "Longitude";

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                int i = 1;
                foreach(var item in pincode)
                {
                    //Add some text to cell A1
                    //worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                    //You could also use [line, column] notation:\

                    if (item.PinCode != "Pincode" && item.District != "District" && item.Latitude != "Latitude" && item.Longitude != "Longitude"){
                        worksheet.Cells[i, 1].Value = item.PinCode;
                        worksheet.Cells[i, 2].Value = item.District;
                        worksheet.Cells[i, 3].Value = item.Latitude;
                        worksheet.Cells[i, 4].Value = item.Longitude;
                    }
                    else
                    {
                        i = i-1;
                    }

                    i += 1;
                }                

                DirectoryInfo directory = new("Pincode.xlsx");
                if (directory.Exists) directory.Delete();
                //Save your file
                FileInfo fi = new FileInfo(@"Pincode.xlsx");
                excelPackage.SaveAs(fi);
            }
        }

    }    
}
