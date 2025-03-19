using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using WebApplicationDocs.Models;


namespace WebApplicationDocs.Controllers
{
    public class FileProcessingController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ProcessFile(FileProcessingModel model)
        {
            if (!ModelState.IsValid)
            {
                return View("Index", model);
            }

            try
            {
                string excelPath = "E:\\Automation\\Excel\\Demo Tin Numbers.xlsx";
                Directory.CreateDirectory(model.DestPath);

                string[] zipFiles = Directory.GetFiles(model.SourcePath, $"{model.ClientId}*.zip");
                if (zipFiles.Length == 0)
                {
                    ViewBag.Message = $"Files for {model.ClientId} not found.";
                    return View("Index", model);
                }

                List<string> tinNumbers = ReadTinNumbersFromExcel(excelPath);
                if (tinNumbers.Count == 0)
                {
                    ViewBag.Message = "No TIN Numbers found.";
                    return View("Index", model);
                }

                string validFile = null;
                foreach (var file in zipFiles.OrderByDescending(f => System.IO.File.GetCreationTime(f)))
                {
                    string copiedFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(file));
                    System.IO.File.Copy(file, copiedFile, true);

                    string unzipDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                    Directory.CreateDirectory(unzipDir);
                    ZipFile.ExtractToDirectory(copiedFile, unzipDir);

                    string[] unZippedFiles = Directory.GetFiles(unzipDir, "*.*");
                    string[] lines = System.IO.File.ReadAllLines(unZippedFiles[0]);
                    string[] firstLineContent = lines[0].Split("\t");
                    string filesDocumentType = firstLineContent[22];

                    if (filesDocumentType == model.DocumentType)
                    {
                        validFile = unZippedFiles[0];
                        ProcessFile(validFile, tinNumbers, unzipDir, model.DestPath, model.PaymentFileNum,model.ReplacementSuffix);
                        break;
                    }
                    Directory.Delete(unzipDir, true);
                }

                if (validFile == null)
                {
                    ViewBag.Message = $"{model.ClientId} Files with document type {model.DocumentType} not found.";
                }
                else
                {
                    ViewBag.Message = "File processed successfully!";
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
            }

            return View("Index", model);
        }

        private void ProcessFile(string filePath, List<string> tinNumbers, string unzipDir, string destPath, int paymentFileNum, string replacementSuffix)
        {
            string[] lines = System.IO.File.ReadAllLines(filePath);
            string currentDate = DateTime.Now.ToString("yyyyMMdd");
            //string replacementSuffix = "P3";

            List<string> updatedLines = new List<string>();
            string newText = "";
            int tinIndex = -1;
            int paymentFileCount = 0;
            bool eraseMode = false;

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("00"))
                {
                    if (paymentFileCount >= tinNumbers.Count || paymentFileCount >= paymentFileNum)
                    {
                        eraseMode = true;
                    }
                    else
                    {
                        paymentFileCount++;
                        tinIndex++;
                    }
                }
                if (eraseMode)
                {
                    break;
                }

                if (lines[i].StartsWith("01") && tinIndex < tinNumbers.Count)
                {
                    string[] fields = lines[i].Split('\t');
                    string oldTIN = fields[55];
                    if (string.IsNullOrEmpty(oldTIN))
                    {
                        return;
                    }
                    lines[i] = lines[i].Replace(oldTIN, tinNumbers[tinIndex]);
                }

                if (lines[i].Length >= 25)
                {
                    string oldText = lines[i].Substring(6, 19);
                    newText = oldText.Substring(0, 9) + currentDate + replacementSuffix;
                    lines[i] = lines[i].Replace(oldText, newText);
                }

                updatedLines.Add(lines[i]);
            }

            System.IO.File.WriteAllLines(filePath, updatedLines, System.Text.Encoding.UTF8);
            string renamedFile = Path.Combine(unzipDir, newText + ".docs");
            System.IO.File.Move(filePath, renamedFile);

            string newZipFile = Path.Combine(destPath, Path.GetFileNameWithoutExtension(renamedFile) + ".zip");
            ZipFile.CreateFromDirectory(unzipDir, newZipFile);

            Directory.Delete(unzipDir, true);
        }

        private List<string> ReadTinNumbersFromExcel(string excelPath)
        {
            List<string> tinNumbers = new List<string>();

            if (!System.IO.File.Exists(excelPath)) return tinNumbers;

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new System.IO.FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string tin = worksheet.Cells[row, 1].Text.Trim();
                    if (!string.IsNullOrEmpty(tin))
                        tinNumbers.Add(tin);
                }
            }

            return tinNumbers;
        }
    }
}
