using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO.Compression;
using System.Reflection;
using WebApplicationDocs.Models;



namespace WebApplicationDocs.Controllers
{
    public class SFinalFileProcessingController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        List<string> updatedLines = new List<string>();
        String newFileName = "";
        int counter = 1;
        int userDemandedPaymentFileCount;
        private static string baseDirectory = @"C:\TrackedFiles"; // Your folder
        private static string logFileName = $"UsedFiles_{DateTime.Now:yyyy-MM-dd}.xlsx";
        private static string logFilePath = System.IO.Path.Combine(baseDirectory, logFileName);
        [HttpPost]
        public ActionResult ProcessFile(SFinalFileProcessingModel model)
        {
            if (!ModelState.IsValid)
            {
                TempData["Message"] = "Provided input is invalid.";  //ViewBag.Message
                return RedirectToAction("Index");//View("Index", model);
            }

            try
            {
                string excelPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), model.ExcelFile.FileName);
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
                userDemandedPaymentFileCount = model.PaymentFileNum;
                string validFile = "";
                int paymentFileNumber = 0;
                //model.PaymentFileNum;
                HashSet<string> usedFiles = LoadUsedFiles();

                foreach (var file in zipFiles.OrderByDescending(f => System.IO.File.GetCreationTime(f)))
                {
                    string copiedFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetFileName(file));
                    System.IO.File.Copy(file, copiedFile, true);

                    string unzipDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName());
                    Directory.CreateDirectory(unzipDir);
                    ZipFile.ExtractToDirectory(copiedFile, unzipDir);

                    string[] unZippedFiles = Directory.GetFiles(unzipDir, "*.*");
                    string[] lines = System.IO.File.ReadAllLines(unZippedFiles[0]);  //008
                    string[] firstLineContent = lines[0].Split("\t");
                    string filesDocumentType = firstLineContent[22];
                    string oldText = lines[0].Substring(6, 19);//old filename
                    if (newFileName.Equals(""))
                    {

                        string currentDate = DateTime.Now.ToString("yyyyMMdd");
                        newFileName = oldText.Substring(0, 9) + currentDate + model.ReplacementSuffix;
                    }
                    //code for checking in Excel if not in Excel the Adding in else part
                    if (usedFiles.Contains(oldText))
                    {
                        Console.WriteLine($"Skipping: {oldText} already used today.");
                        continue;
                    }
                    else
                    {
                        SaveFileToLog(oldText);
                    }

                    int validPaymentCount = 0;

                    int FilePaymentCount = int.Parse(lines[lines.Length - 1].Split("\t")[2].Substring(19)); //total num of payment files
                    Dictionary<int, int> paymentFileCountWithIndex = new Dictionary<int, int>();
                    Dictionary<int, int[]> validPaymentFileCountWithIndex = new Dictionary<int, int[]>();
                    int actualValidPaymentFiles = 0;
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].StartsWith("00"))
                        {
                            paymentFileCountWithIndex.Add(validPaymentCount++, i);
                        }
                    }
                    for (int i = 0; i < validPaymentCount; i++)
                    {
                        Boolean isDocumentTypeIsValid = false;
                        int firstindex = paymentFileCountWithIndex.GetValueOrDefault(i);
                        string[] firstindexLineContent = lines[firstindex].Split("\t");
                        string paymentfileDocumentType = firstLineContent[22];
                        if (model.DocumentType == paymentfileDocumentType)
                        {
                            isDocumentTypeIsValid = true;
                        }
                        else
                        {
                            break;
                        }
                        int lastindex = 0;
                        if (i == validPaymentCount)
                        {
                            lastindex = paymentFileCountWithIndex.GetValueOrDefault(lines.Length) - 1;
                        }
                        else
                        {
                            lastindex = paymentFileCountWithIndex.GetValueOrDefault(i + 1) - 1;
                        }
                        Boolean isPaymentFileContains43 = false;
                        for (int j = firstindex; j <= lastindex; j++)
                        {
                            if (lines[j].StartsWith("43"))
                            {
                                isPaymentFileContains43 = true;
                            }
                        }

                        if (isPaymentFileContains43 == true && isDocumentTypeIsValid == true)
                        {
                            //copy or counter 
                            int[] firstAndLastIndex = { firstindex, lastindex };
                            validPaymentFileCountWithIndex.Add(actualValidPaymentFiles++, firstAndLastIndex);
                        }
                    }
                    if (validPaymentFileCountWithIndex.Count > 0)  //10  //43
                    {
                        validFile = unZippedFiles[0];
                        //ProcessFile(validPaymentFileCountWithIndex, actualValidPaymentFiles, validFile, tinNumbers, unzipDir, model.DestPath, model.PaymentFileNum, model.ReplacementSuffix);
                        UpdateListOfLines(updatedLines, validPaymentFileCountWithIndex, actualValidPaymentFiles, validFile, tinNumbers, unzipDir, model.DestPath, model.PaymentFileNum, model.ReplacementSuffix);
                        System.Diagnostics.Debug.WriteLine("validPaymentFileCountWithIndex >>>>>" + validPaymentFileCountWithIndex.Count);
                        if (userDemandedPaymentFileCount == (counter - 1))
                        {
                            //add to file
                            System.IO.File.WriteAllLines(validFile, updatedLines, System.Text.Encoding.ASCII);
                            string renamedFile = System.IO.Path.Combine(unzipDir, newFileName + ".docs");
                            System.IO.File.Move(validFile, renamedFile);

                            string newZipFile = System.IO.Path.Combine(model.DestPath, System.IO.Path.GetFileNameWithoutExtension(renamedFile) + ".zip");
                            ZipFile.CreateFromDirectory(unzipDir, newZipFile);

                            Directory.Delete(unzipDir, true);
                            TempData["Message"] = "File processed successfully!";
                            break;
                        }
                    }

                }

                //create the file and add the lines 

                if (validFile == null)
                {
                    TempData["Message"] = $"{model.ClientId} Files with document type {model.DocumentType} not found.";
                }
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }

            return View("Index", model);
        }

        private void UpdateListOfLines(List<string> updatedLines, Dictionary<int, int[]> validPaymentFileCountWithIndex, int actualValidPaymentFiles, string filePath, List<string> tinNumbers, string unzipDir, string destPath, int paymentFileNum, string replacementSuffix)
        {
            try
            {
                string[] lines = System.IO.File.ReadAllLines(filePath);
                string currentDate = DateTime.Now.ToString("yyyyMMdd");
                string newText = "";
                int tinIndex = 0;

                if (tinNumbers.Count < paymentFileNum)
                {
                    TempData["Message"] = "Provided Excel file has TIn Numbers less than the given paymentFileNum ";
                    return;
                }

                for (int i = 0; i <= actualValidPaymentFiles; i++)
                {
                    if (userDemandedPaymentFileCount == (counter - 1))
                    {
                        System.Diagnostics.Debug.WriteLine("Counter Succesfully completed as per the user expected Payment File count");
                        break;
                    }
                    int start = 0, end = 0;
                    if (validPaymentFileCountWithIndex != null)
                    {
                        int[] firstAndLastIndex = validPaymentFileCountWithIndex.GetValueOrDefault(i);
                        if (firstAndLastIndex != null)
                        {
                            start = firstAndLastIndex[0];
                            end = firstAndLastIndex[1];
                        }
                    }
                    string oldTIN = string.Empty;
                    for (int j = start; j <= end; j++)
                    {
                        if (lines[j].StartsWith("01"))
                        {
                            string[] fields = lines[j].Split('\t');
                             oldTIN = fields[55];    
                        }
                        if (!string.IsNullOrEmpty(oldTIN))
                        {
                            lines[j] = lines[j].Replace(oldTIN, tinNumbers[tinIndex]);
                            Console.WriteLine("New Tin number at" + tinIndex + ">>>>" + tinNumbers[tinIndex]);
                        }
                            

                        string oldText = lines[j].Substring(6, 25);
                        string paymentFileNumber = counter.ToString("D6");

                        Console.WriteLine("paymentFileNumber is" + paymentFileNumber);
                        newText = newFileName + paymentFileNumber;
                        Console.WriteLine("newText is" + newText);
                        lines[j] = lines[j].Replace(oldText, newText);
                        updatedLines.Add(lines[j]);
                    }
                    counter++;
                    tinIndex++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
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