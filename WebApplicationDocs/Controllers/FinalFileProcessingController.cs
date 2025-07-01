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

                string validFile = null;
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

                    int[] desiredPayments = new int[0];
                    int FilePaymentCount = int.Parse(lines[lines.Length - 1].Split("\t")[2].Substring(19)); //First condition (write this at last)

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
                            string[] currentDocumenttype = lines[i].Split("\t");
                            if (currentDocumenttype.Length > 22)
                            {
                                if (currentDocumenttype[22] == model.DocumentType)
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
            //add a condition to choose only the required payment files
            
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
                        lines[i] = lines[i].Replace(oldTIN, tinNumbers[tinIndex]); //replace all TIN Numbers in the current payment file
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
    }
}
