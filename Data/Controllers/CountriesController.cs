using ImportData.Data;
using ImportData.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ImportData.Controllers
{
   //[ApiController]
    //[Route("api/[controller]")]
    public class HomeController : Controller
    {

        private readonly CountriesAPIDbContext dbContext;

        public HomeController(CountriesAPIDbContext dbContext)
        {
            this.dbContext = dbContext;
        }

        [HttpGet]
        public async Task<IActionResult> GetCountries()
        {
            var countries = await dbContext.Countries.ToListAsync();
            return Ok(countries);
        }

        //[HttpGet]
        //[Route("ExportToExcel")]
        public async Task<IActionResult> ExportToExcel()
        {
            var data = await dbContext.Countries.ToListAsync();

            // Create a new Excel package
            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the Excel package
                var worksheet = package.Workbook.Worksheets.Add("Countries");

                // Set the column headers
                worksheet.Cells[1, 1].Value = "Country ID";
                worksheet.Cells[1, 2].Value = "Country Name";
                worksheet.Cells[1, 3].Value = "Two Char Country Code";
                worksheet.Cells[1, 4].Value = "Three Char Country Code";

                int row = 2;
                foreach (var country in data)
                {
                    worksheet.Cells[row, 1].Value = country.CountryID;
                    worksheet.Cells[row, 2].Value = country.CountryName;
                    worksheet.Cells[row, 3].Value = country.TwoCharCountryCode;
                    worksheet.Cells[row, 4].Value = country.ThreeCharCountryCode;
                    row++;
                }

                // Export the Excel file and return it as the response
                // Auto-fit the columns
                worksheet.Cells.AutoFitColumns();

                // Convert the Excel package to a byte array
                var excelBytes = package.GetAsByteArray();

                // Set the response content type and headers
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.Headers.Add("Content-Disposition", "attachment; filename=Countries.xlsx");

                // Write the Excel data to the response stream
                await Response.Body.WriteAsync(excelBytes);

                return Ok();
            }
        }

        public IActionResult Index()
        {
            ViewData["Title"] = "Home Page";
            return View();
        }




        //private List<Countries> countries = new List<Countries>
        //{
        //    new Countries{ CountryID="11" ,CountryName="India", TwoCharCountryCode="IN",ThreeCharCountryCode="IND"},
        //    new Countries{ CountryID="12" ,CountryName="Australia", TwoCharCountryCode="AU",ThreeCharCountryCode="AUS"}
        //};


        //public IActionResult ExportToCSV(Countries countries)
        //{
        //    var builder = new StringBuilder();
        //    builder.AppendLine(" CountryID ,CountryName, TwoCharCountryCode,ThreeCharCountryCode");
        //    foreach (var country in Countries)
        //    {
        //        builder.AppendLine($"{country.CountryID},{country.CountryName},{country.TwoCharCountryCode},{country.ThreeCharCountryCode}");
        //    }
        //    return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "countries.csv");
        //}
        //public IActionResult Index()
        //{
        //    return View();
        //}


        [HttpPost("ImportCountries")]
        public async Task<IActionResult> Import(IFormFile file)
        {
            try
            {
                var countries = await ImportfromExcel(file);
                dbContext.Countries.AddRange(countries);
                await dbContext.SaveChangesAsync();

                return Ok("Data imported successfully");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error: {ex.Message}");
            }
        }
        public async Task<List<Countries>> ImportfromExcel(IFormFile file)
        {
            var list = new List<Countries>();
            using (MemoryStream stream = new MemoryStream())

            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowcount = worksheet.Dimension.Rows;
                    if (rowcount > 0)
                    {
                        for (int row = 2; row <= rowcount; row++)
                        {
                            list.Add(new Countries
                            {
                                CountryID = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                CountryName = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                TwoCharCountryCode = worksheet.Cells[row, 3].Value.ToString().Trim(),
                                ThreeCharCountryCode = worksheet.Cells[row, 4].Value.ToString().Trim()
                            });
                        }
                    }
                    else
                    {

                        Console.WriteLine("The Excel sheet is empty.");

                    }


                }



            }
            return list;


        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
