﻿using DocumentFormat.OpenXml.EMMA;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO.Compression;

namespace ExcelSplit.Controllers
{
    
        [ApiController]
        [Route("[controller]")]
        public class ExcelSplitController : ControllerBase
        {
        [HttpPost("upload")]
        public async Task<IActionResult> Upload([FromForm] ExcelUpload excelUpload)
        {
            if (excelUpload.FormFile == null || excelUpload.FormFile.Length == 0)
                return BadRequest("No file uploaded.");

            var extension = Path.GetExtension(excelUpload.FormFile.FileName);
            if (extension != ".xls" && extension != ".xlsx")
                return BadRequest("Invalid file type. Please upload an Excel file.");

            var filePath = Path.GetTempFileName();
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await excelUpload.FormFile.CopyToAsync(stream);
            }

            // Process the Excel file

            var outputDirectory = Path.Combine(Path.GetTempPath(), "SplitExcel");
            Directory.CreateDirectory(outputDirectory);

            using (var file = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = file.Workbook; // Get the first worksheet
                var worksheetCount = workbook.Worksheets.Count;

               if (worksheetCount < 1) 
                    return BadRequest("Empty File");

                if (worksheetCount == 1)
                    return Content("File has single worksheet");

                for (int i = 0; i < worksheetCount; i++)
                {
                    var worksheet = file.Workbook.Worksheets[i];
                    var newFile = new ExcelPackage();
                    var newWorksheet = newFile.Workbook.Worksheets.Add(worksheet.Name);

                    // Copy all cells from the original worksheet to the new worksheet
                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;
                    worksheet.Cells[1, 1, rowCount, colCount].Copy(newWorksheet.Cells[1, 1]);

                    // Save the new workbook
                    var newFileName = Path.Combine(outputDirectory, $"{worksheet.Name}.xlsx");
                    newFile.SaveAs(new FileInfo(newFileName));
                }

            }

            var zipPath = Path.Combine(Path.GetTempPath(), "SplitWorkbooks.zip");
            ZipFile.CreateFromDirectory(outputDirectory, zipPath);

            // Return the ZIP file as a download
            var zipFileStream = new FileStream(zipPath, FileMode.Open);
            var zipFile = new FileStreamResult(zipFileStream, "application/zip")
            {
                FileDownloadName = "SplitWorkbooks.zip"
            };

            //return Ok("File uploaded and processed successfully.");
            return zipFile;
        }
    }
}
