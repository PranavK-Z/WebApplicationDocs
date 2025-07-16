using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using WebApplicationDocs.Models;

namespace WebApplicationDocs.Controllers
{
    public class FindTheFileController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ProcessFile(FindTheFileModel model)
        {
            if (!ModelState.IsValid)
            {
                TempData["Message"] = "Provided input is invalid.";
                return RedirectToAction("Index");
            }

            try
            {
                
                Directory.CreateDirectory(model.DestPath);

                string[] zipFiles = Directory.GetFiles(model.SourcePath, $"{model.ClientId}*.zip");//client... .zip//change
                if (zipFiles.Length == 0)
                {
                    TempData["Message"] = $"Files for {model.ClientId} not found.";
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
                    string[] lines = System.IO.File.ReadAllLines(unZippedFiles[0]);
                    string[] firstLineContent = lines[0].Split("\t");
                    string filesDocumentType = firstLineContent[22];
                    String fileRecipientType = firstLineContent[21];

                    int FilePaymentCount = int.Parse(lines[lines.Length - 1].Split("\t")[2].Substring(19));

                    if (filesDocumentType == model.DocumentType && fileRecipientType == model.RecipientType && FilePaymentCount + 1 >= model.PaymentFileNum)
                    {
                        validFile = unZippedFiles[0];
                        
                        string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                        Directory.CreateDirectory(tempDir);

                        // Copy the valid file to the temporary directory
                        string tempFilePath = Path.Combine(tempDir, Path.GetFileName(validFile));
                        System.IO.File.Copy(validFile, tempFilePath, true);

                        // Create the zip file from the temporary directory
                        string destinationFilePath = Path.Combine(model.DestPath, Path.GetFileName(validFile) + ".zip");
                        destinationFilePath = destinationFilePath.Replace(".Docs", "");
                        ZipFile.CreateFromDirectory(tempDir, destinationFilePath);

                        // Clean up the temporary directory
                        Directory.Delete(tempDir, true);
                        TempData["Message"] = $"{model.ClientId} Files with document type {model.DocumentType} and number of payment files more than {model.PaymentFileNum} is pasted at {model.DestPath}.";
                        break;
                        //1.Path.GetFileName(validFile): Extracts the file name(e.g., example.txt) from the validFile path.
                        //2.Path.Combine(model.DestPath, ...): Combines the destination directory(model.DestPath) with the file name to create the full destination file path.
                        //3.System.IO.File.Copy: Copies the file from validFile(source) to the constructed destinationFilePath.
                        //ProcessFile(validFile, tinNumbers, unzipDir, model.DestPath, model.PaymentFileNum, model.ReplacementSuffix);

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

    }
}
