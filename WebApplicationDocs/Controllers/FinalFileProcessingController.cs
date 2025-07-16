using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using WebApplicationDocs.Models;

namespace WebApplicationDocs.Controllers
{
    public class FinalFileProcessingController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private static string baseDirectory = @"C:\TrackedFiles"; // Your folder
        private static string logFileName = $"UsedFiles_{DateTime.Now:yyyy-MM-dd}.xlsx";
        private static string logFilePath = System.IO.Path.Combine(baseDirectory, logFileName);

        [HttpPost]
        public ActionResult ProcessFile(FinalFileProcessingModel model)
        {
            if (!ModelState.IsValid)
            {
                TempData["Message"] = "Provided input is invalid.";  //ViewBag.Message
                return RedirectToAction("Index");//View("Index", model);
            }

            try
            {
                string excelPath = Path.Combine(Path.GetTempPath(), model.ExcelFile.FileName);
                using (var stream = new FileStream(excelPath, FileMode.Create))
                {
                    model.ExcelFile.CopyTo(stream);
                }

                if (!System.IO.File.Exists(excelPath))
                {
                    TempData["Message"] = "Invalid Excel file.";
                    return RedirectToAction("Index");
                }

                Directory.CreateDirectory(model.DestPath);

                string[] zipFiles = Directory.GetFiles(model.SourcePath, $"{model.ClientId}*.zip");
                if (zipFiles.Length == 0)
                {
                    TempData["Message"] = $"Files for {model.ClientId} not found.";
                    return RedirectToAction("Index"); 
                }

                List<string> tinNumbers = ReadTinNumbersFromExcel(excelPath);
                if (tinNumbers.Count == 0)
                {
                    TempData["Message"] = "No TIN Numbers found.";
                    return RedirectToAction("Index");
                }

                if (tinNumbers.Count < model.PaymentFileNum)////
                {
                    TempData["Message"] = "Please add Sufficient TIN numbers as per the payment file requested";
                    return RedirectToAction("Index");
                }

                string validFile = null;
                HashSet<string> usedFiles = LoadUsedFiles();

                foreach (var file in zipFiles.OrderByDescending(f => System.IO.File.GetCreationTime(f)))
                {
                    string copiedFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(file));
                    System.IO.File.Copy(file, copiedFile, true);

                    string unzipDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                    Directory.CreateDirectory(unzipDir);
                    ZipFile.ExtractToDirectory(copiedFile, unzipDir);

                    string[] unZippedFiles = Directory.GetFiles(unzipDir, "*.*");
                    string[] lines = System.IO.File.ReadAllLines(unZippedFiles[0]); //choose .docs file
                    //string[] firstLineContent = lines[0].Split("\t");
                    //string filesDocumentType = firstLineContent[22]; //
                    string oldText = lines[0].Substring(6, 19);
                    int[] desiredPayments = new int[0];
                    int FilePaymentCount = int.Parse(lines[lines.Length - 1].Split("\t")[2].Substring(19)); 

                    if (usedFiles.Contains(oldText))
                    {
                        Console.WriteLine($"Skipping: {oldText} already used today.");
                        continue;
                    }
                    //First condition (write this at last)
                    if (FilePaymentCount < model.PaymentFileNum)//add documenttype condition too
                    {
                        continue; //Skip this file if the payment file count is less than the required payment file number
                    }
                    //code for checking in Excel if not in Excel the Adding in else part
                    
                    //else
                    //{
                    //    SaveFileToLog(oldText);
                    //}
                    //second condition: Counting desired payment files and checking if it is greater than or equal to the given paymentFileNum
                    //simuttaneousy noting the valid payment file numbers
                    Boolean isDocumentTypeMatch = false;
                    int temp=0;
                    for (int i = 0; i < lines.Length; i++)
                    {
                        int  currentPayment = int.Parse(lines[i].Split("\t")[2].Substring(19));
                        
                        if (currentPayment > temp)  //to reset DocumentType for each payment file.
                        {
                            isDocumentTypeMatch = false;
                        }
                        
                        Boolean hasRecord43 = false;
                        if (lines[i].StartsWith("00") )
                        {
                            string[] currentDocumenttypeRecipientType = lines[i].Split("\t");
                            if (currentDocumenttypeRecipientType.Length > 22)
                            {
                                if (currentDocumenttypeRecipientType[22] == model.DocumentType && currentDocumenttypeRecipientType[21] == model.RecipientType)
                                {
                                    isDocumentTypeMatch = true;
                                }
                            }    
                        }
                        if(lines[i].StartsWith("43")) //multiple 43records
                        {
                            hasRecord43 = true;
                        }
                        Boolean existingPayment = false;  //If a payment files has multiple 43 records this will be used to avoid duplicates
                        if (isDocumentTypeMatch == true && hasRecord43 ==true)
                        {
                            for(int j = 0; j < desiredPayments.Length; j++)
                            {
                                if (desiredPayments[j] == currentPayment)
                                {
                                    existingPayment = true;
                                   //break;
                                }
                            }
                            if(existingPayment == false) {
                                desiredPayments = desiredPayments.Append(currentPayment).ToArray();
                            }
                            
                        }
                        temp = currentPayment;
                    }

                    //if (filesDocumentType == model.DocumentType && FilePaymentCount + 1 >= model.PaymentFileNum)
                    if ( desiredPayments.Length >= model.PaymentFileNum) //filesDocumentType == model.DocumentType &&
                    {
                        validFile = unZippedFiles[0];
                        ProcessFile(validFile, tinNumbers, unzipDir, model.DestPath, model.PaymentFileNum, model.ReplacementSuffix,desiredPayments);
                        SaveFileToLog(oldText);
                        break;
                    }
                    Directory.Delete(unzipDir, true);
                }

                if (validFile == null)
                {
                    TempData["Message"] = $"{model.ClientId} Files with document type {model.DocumentType} not found.";
                }

            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }

            return RedirectToAction("Index");
        }

        private void ProcessFile(string filePath, List<string> tinNumbers, string unzipDir, string destPath, int paymentFileNum, string replacementSuffix, Array requiredPayments)
        {
            string[] lines = System.IO.File.ReadAllLines(filePath);
            string currentDate = DateTime.Now.ToString("yyyyMMdd");

            List<string> updatedLines = new List<string>();
            string newText = "";
            int tinIndex = -1;
            int paymentFileCount = 0;
            bool eraseMode = false;

            if (tinNumbers.Count < paymentFileNum)
            {
                TempData["Message"] = "Provided Excel file has TIn Numbers less than the given paymentFileNum ";
                return;
            }
            
            string oldTIN = string.Empty;
            for (int i = 0; i < lines.Length; i++)
            {
                int currentPayment = int.Parse(lines[i].Split("\t")[2].Substring(19));
                //if this current payment is in the required payments, then process it
                Boolean process = false;
                for (int j = 0; j < requiredPayments.Length; j++)
                {
                    if (currentPayment == (int)requiredPayments.GetValue(j))
                    {
                        process = true;
                        break;
                    }
                }
                
                if (process ==true)
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
                            tinIndex++;//
                        }
                    }
                    if (eraseMode)
                    {
                        break;
                    }
                   
                    if (lines[i].StartsWith("01"))
                    {
                        string[] fields = lines[i].Split('\t');
                        oldTIN = fields[55];
                        if (string.IsNullOrEmpty(oldTIN))
                        {
                            throw new InvalidOperationException("TIN number not found.");   
                        }
                    }

                    if (!lines[i].StartsWith("00") && tinIndex < tinNumbers.Count)
                    {
                        if(!string.IsNullOrEmpty(oldTIN))
                        {
                            lines[i] = lines[i].Replace(oldTIN, tinNumbers[tinIndex]); //oldTIN ni tinNumbers[] tho replace cheyali 
                        }      
                    }

                    if (lines[i].Length >= 25)
                    {
                        string oldText = lines[i].Substring(6, 19);
                        newText = oldText.Substring(0, 9) + currentDate + replacementSuffix;
                        lines[i] = lines[i].Replace(oldText, newText);
                    }

                    updatedLines.Add(lines[i]);
                }
               
            }

            System.IO.File.WriteAllLines(filePath, updatedLines, System.Text.Encoding.ASCII);
            string renamedFile = Path.Combine(unzipDir, newText + ".docs");
            System.IO.File.Move(filePath, renamedFile);

            string newZipFile = Path.Combine(destPath, Path.GetFileNameWithoutExtension(renamedFile) + ".zip");
            ZipFile.CreateFromDirectory(unzipDir, newZipFile);

            Directory.Delete(unzipDir, true);
            TempData["Message"] = "File processed successfully!";
        }

        private List<string> ReadTinNumbersFromExcel(string excelPath)
        {
            List<string> tinNumbers = new List<string>();

            try
            {
                if (!System.IO.File.Exists(excelPath)) return tinNumbers;

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new System.IO.FileInfo(excelPath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    if (worksheet.Dimension == null) return tinNumbers;
                    int rowCount = worksheet.Dimension.Rows;


                    for (int row = 1; row <= rowCount; row++)
                    {
                        string tin = worksheet.Cells[row, 1].Text.Trim();
                        if (!string.IsNullOrEmpty(tin))
                            tinNumbers.Add(tin);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error reading Excel file: " + ex.Message;
            }

            return tinNumbers;
        }

        private static HashSet<string> LoadUsedFiles()
        {
            var usedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!System.IO.File.Exists(logFilePath))
                return usedFiles;

            using (var workbook = new XLWorkbook(logFilePath))
            {
                var ws = workbook.Worksheet(1);
                foreach (var row in ws.RowsUsed())
                {
                    var value = row.Cell(1).GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                        usedFiles.Add(value.Trim());
                }
            }

            return usedFiles;
        }

        private static void SaveFileToLog(string fileName)
        {
            XLWorkbook workbook;
            IXLWorksheet worksheet;

            if (System.IO.File.Exists(logFilePath))
            {
                workbook = new XLWorkbook(logFilePath);
                worksheet = workbook.Worksheet(1);
            }
            else
            {
                Directory.CreateDirectory(baseDirectory);
                workbook = new XLWorkbook();
                worksheet = workbook.AddWorksheet("UsedFiles");
                worksheet.Cell(1, 1).Value = "FileName";
            }

            // Add to next available row
            int lastRow = worksheet.LastRowUsed().RowNumber();
            worksheet.Cell(lastRow + 1, 1).Value = fileName;

            workbook.SaveAs(logFilePath);
        }

    }
}
