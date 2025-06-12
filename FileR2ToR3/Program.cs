using DbR2;
using DbR3;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
namespace FileR2ToR3
{
    class Program
    {
        public static DbR2.DBR2DataContext dbR2 = new DbR2.DBR2DataContext(ConfigurationSettings.AppSettings["DbR2"]);
        public static DbR3.DBR3DataContext dbR3 = new DbR3.DBR3DataContext(ConfigurationSettings.AppSettings["DbR3"]);
        public static string conR2 = ConfigurationSettings.AppSettings["DbR2"];
        public static string conR3 = ConfigurationSettings.AppSettings["DbR3"];
        public static string filePathR2 = ConfigurationSettings.AppSettings["TempPathR2"];
        public static string filePathR3 = ConfigurationSettings.AppSettings["TempPathR3"];
        public static string user = "";
        public static string pass = "";
        
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            try
            {
                user = ConfigurationSettings.AppSettings["UserPass"].Split('|').ElementAtOrDefault(0);
                pass = ConfigurationSettings.AppSettings["UserPass"].Split('|').ElementAtOrDefault(1);
                filePathR2 = CleanTempPath(filePathR2);
                filePathR3 = CleanTempPath(filePathR3);
                Console.WriteLine("Welcome To Program : FileR2ToR3");
                Console.WriteLine($"     filePathR2 : {filePathR2}");
                Console.WriteLine($"     ConnectionString R2 : {conR2}");
                Console.WriteLine($"     filePathR3 : {filePathR3}");
                Console.WriteLine($"     ConnectionString R3 : {conR3}");
                Console.WriteLine("This program will convert a file from R2 to R3 format.");
                Console.WriteLine("Please select the mode of attachment you want to convert");

                Console.WriteLine("1. Attachment");
                Console.WriteLine("2. TRNDelegate & TRNDelegateDetail");
                Console.WriteLine("3. DelegateAttachment");
                Console.WriteLine("4. ControlAttachment(in memo detail)");
                Console.WriteLine("5. Convert SignaturePath(Base64 To TempFile)");
                Console.WriteLine("6. Convert CompanyLogo(Base64 To TempFile)");

                Console.WriteLine("Please enter Mode : ");
                var mode = Convert.ToInt32(Console.ReadLine());

                Console.Clear();
                switch (mode)
                {
                    case 1:
                        Attachment();
                        break;
                    case 2:
                        DeleagteData();
                        break;
                    case 3:
                        break;
                    case 4:
                        break;
                    case 5:
                        EmployeeConvertBase64ToTempFile();
                        break;
                    case 6:
                        CompanyConvertBase64ToTempFile();
                        break;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            
            Thread.Sleep(1000000000);
        }

        public static void DeleagteData() 
        {
            var delegates = dbR2.TRNDelegates.ToList();
            int total = delegates.Count;
            int current = 0;
            foreach (var delegator in delegates)
            {
                current++;
                var appCodeR2 = dbR2.MSTEmployees.FirstOrDefault(x => x.EmployeeId == delegator.ApproverId)?.EmployeeCode;
                var deleCodeR2 = dbR2.MSTEmployees.FirstOrDefault(x => x.EmployeeId == delegator.DelegateToId)?.EmployeeCode;

                var newDelegate = new DbR3.TRNDelegate
                {
                    ApproverId = dbR3.MSTEmployees.FirstOrDefault(x => x.EmployeeCode == appCodeR2)?.EmployeeId,
                    DelegateToId = dbR3.MSTEmployees.FirstOrDefault(x => x.EmployeeCode == deleCodeR2)?.EmployeeId,
                    DateFrom = delegator.DateFrom,
                    DateTo = delegator.DateTo,
                    CreatedDate = DateTime.Now,
                    Remark = delegator.Remark,
                    IsActive = delegator.IsActive,
                    AccountId = delegator.AccountId,
                    DelegateToRole = delegator.DelegateToRole,
                    IsApplySomeForm = delegator.IsApplySomeForm,
                };
                dbR3.TRNDelegates.InsertOnSubmit(newDelegate);
                dbR3.SubmitChanges();
                // Copy TRNDelegateDetails
                var deleDetails = dbR2.TRNDelegateDetails.Where(x => x.DelegateId == delegator.DelegateId).ToList();
                foreach (var dele in deleDetails)
                {
                    var memoDoc = dbR2.TRNMemos.FirstOrDefault(x => x.MemoId == dele.MemoId)?.DocumentNo;
                    var memoIdR3 = dbR3.TRNMemos.FirstOrDefault(x => x.DocumentNo == memoDoc)?.MemoId;
                    var temp = dbR2.MSTTemplates.FirstOrDefault(x => x.TemplateId == dele.TemplateId)?.DocumentCode;
                    var tempR3 = dbR3.MSTTemplates.FirstOrDefault(x => x.TemplateCode == temp)?.TemplateId;
                    var newDelegateDetail = new DbR3.TRNDelegateDetail
                    {
                        DelegateId = newDelegate.DelegateId,
                        CreatedDate = DateTime.Now,
                        IsActive = dele.IsActive,
                        MemoId = (int?)memoIdR3,
                        TemplateId = tempR3,
                    };
                    dbR3.TRNDelegateDetails.InsertOnSubmit(newDelegateDetail);
                }
                ShowProgress(current, total);
            }
        }

        public static void Attachment()
        {
            var attachments = dbR2.TRNAttachFiles.Where(x => x.MemoId != null).ToList();
            int total = attachments.Count;
            int current = 0;
            foreach (var attactment in attachments)
            {
                current++;
                var documentNo = dbR2.TRNMemos.FirstOrDefault(x => x.MemoId == attactment.MemoId)?.DocumentNo;
                if (string.IsNullOrEmpty(documentNo))
                {
                    WriteLog($"[SKIP] DocumentNo not found for MemoId: {attactment.MemoId}");
                    continue;
                }
                var memoIdR3 = dbR3.TRNMemos.FirstOrDefault(x => x.DocumentNo == documentNo)?.MemoId;
                var newFileName = Guid.NewGuid().ToString();
                var description = attactment.FileName;
                var fullFileName = attactment.AttachFile;
                var filePath = CleanTempPath(attactment.FilePath);
                var memoId = attactment.MemoId;
                var ext = Path.GetExtension(fullFileName).ToLowerInvariant();
                if (filePath.Contains("sharepoint.com"))
                {
                }
                else
                {
                    CopyFileWithNewName(filePath, newFileName);
                }
                var newFile = new DbR3.TRNAttachFile
                {
                    FileName = Path.GetFileNameWithoutExtension(fullFileName),
                    FileOriginalName = fullFileName,
                    Description = description,
                    FilePath = filePath.Contains("sharepoint.com") ? attactment.FilePath : newFileName,
                    MemoId = (int?)memoIdR3,
                    MimeTypeId = GetMimeTypeId(ext),
                    CreatedDate = DateTime.Now,
                };
                dbR3.TRNAttachFiles.InsertOnSubmit(newFile);
                ShowProgress(current, total);
                dbR3.SubmitChanges();
            }
        }

        public static void EmployeeConvertBase64ToTempFile()
        {
            var employees = dbR3.MSTEmployees.ToList();
            int total = employees.Count;
            int current = 0;
            foreach (var employee in employees)
            {
                current++;

                if (!string.IsNullOrEmpty(employee.SignPicPath) && employee.SignPicPath.Contains("base64"))
                {
                    var ext = GetImageExtensionFromBase64(employee.SignPicPath);
                    var picBytes = Base64ToMemoryStream(employee.SignPicPath);
                    var newFileName = Guid.NewGuid().ToString();
                    var mimeTypeId = GetMimeTypeId(ext);
                    var fileName = $"signPic_{employee.EmployeeId}{ext}";

                    var newFile = new DbR3.TRNAttachFile
                    {
                        FileName = Path.GetFileNameWithoutExtension(fileName),
                        FileOriginalName = fileName,
                        Description = "",
                        FilePath = newFileName,
                        MimeTypeId = mimeTypeId,
                        CreatedDate = DateTime.Now,
                    };

                    dbR3.TRNAttachFiles.InsertOnSubmit(newFile);
                    SaveFile(picBytes.ToArray(), newFileName);
                    employee.SignPicPath = newFileName;
                }
                else
                {
                    WriteLog($"[SKIP] EmployeeId: {employee.EmployeeId} SignPicPath is null or empty");
                }

                dbR3.SubmitChanges();
                ShowProgress(current, total);
            }
        }
        public static void CompanyConvertBase64ToTempFile()
        {
            var comanies = dbR3.MSTCompanies.ToList();
            int total = comanies.Count;
            int current = 0;
            foreach (var company in comanies)
            {
                current++;

                if (!string.IsNullOrEmpty(company.UrlLogo) && company.UrlLogo.Contains("base64"))
                {
                    var ext = GetImageExtensionFromBase64(company.UrlLogo);
                    var picBytes = Base64ToMemoryStream(company.UrlLogo);
                    var newFileName = Guid.NewGuid().ToString();
                    var mimeTypeId = GetMimeTypeId(ext);
                    var fileName = $"PicLogo_{company.CompanyId}{ext}";

                    var newFile = new DbR3.TRNAttachFile
                    {
                        FileName = Path.GetFileNameWithoutExtension(fileName),
                        FileOriginalName = fileName,
                        Description = "",
                        FilePath = newFileName,
                        MimeTypeId = mimeTypeId,
                        CreatedDate = DateTime.Now,
                    };

                    dbR3.TRNAttachFiles.InsertOnSubmit(newFile);
                    SaveFile(picBytes.ToArray(), newFileName);
                    company.UrlLogo = newFileName;
                }
                else
                {
                    WriteLog($"[SKIP] CompanyId: {company.CompanyId} UrlLogo is null or empty");
                }

                dbR3.SubmitChanges();
                ShowProgress(current, total);
            }
        }
        public static void WriteLog(string message)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "Log");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            var logFilePath = Path.Combine(path, $"log_{DateTime.Now:yyyy-MM-dd}.txt");
            if (!File.Exists(logFilePath))
            {
                File.Create(logFilePath).Close();
            }
            using (var writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
            }

            int progressLine = 0;
            int nextLine = Console.CursorTop;

            if (nextLine <= progressLine)
                nextLine = progressLine + 1;

            Console.SetCursorPosition(0, nextLine);

            // ถ้ามี [SKIP] ให้แยกส่วนสีแดง
            if (message.StartsWith("[SKIP]"))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("[SKIP]");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(message.Substring(6)); // ส่วนที่เหลือ
            }
            else
            {
                Console.WriteLine(message);
            }
        }

        public static void ShowProgress(int current, int total)
        {
            int width = 50;
            double percent = (double)current / total;
            int filled = (int)(percent * width);
            string bar = new string('█', filled) + new string('-', width - filled);
            string text = $"Progress: [{bar}] {percent:P0}";

            int currentLeft = Console.CursorLeft;
            int currentTop = Console.CursorTop;

            Console.SetCursorPosition(0, 0); // ไปแถวบนสุด
            Console.Write(text.PadRight(Console.WindowWidth)); // Clear เหลือบรรทัด
            Console.SetCursorPosition(currentLeft, currentTop); // กลับตำแหน่งเดิม
        }
        public static int GetMimeTypeId(string ext)
        {
            return dbR3.MSTMimeTypes.FirstOrDefault(x => x.Extension == ext)?.MimeTypeId ?? dbR3.MSTMimeTypes.FirstOrDefault(x => x.MimeType == "application/octet-stream").MimeTypeId;
        }

        public static string SaveFile(byte[] fileBytes, string fileName)
        {
            string filePath = Path.Combine(filePathR3, fileName);

            File.WriteAllBytes(filePath, fileBytes);

            return fileName;
        }
        public static string CleanTempPath(string tempPath)
        {
            if (string.IsNullOrWhiteSpace(tempPath))
                WriteLog("Path cannot be null or empty");
            tempPath = tempPath.TrimStart('/', '\\');
            tempPath = tempPath.Replace('/', Path.DirectorySeparatorChar);
            if (tempPath.Contains(".."))
                WriteLog("Invalid path: Path traversal detected");

            return tempPath;
        }

        public static string CopyFileWithNewName(string sourceFileName, string newFileName)
        {
            string sourceFilePath = Path.Combine(filePathR2, sourceFileName);
            string destFilePath = Path.Combine(filePathR3, newFileName);

            if (!File.Exists(sourceFilePath))
            {
                WriteLog($"[SKIP] Source file does not exist: {sourceFilePath}");
                return null;
            }

            File.Copy(sourceFilePath, destFilePath, true);
            return newFileName;
        }

        public static byte[] DownloadFileFromSharePoint(string sharePointUrl, string username, string password, string domain = null)
        {
            // Pseudocode:
            // 1. Create a WebClient instance.
            // 2. If credentials are provided, set them on the WebClient.
            // 3. Download the file as a byte array from the SharePoint
            // URL.
            // 4. Return the byte array.

            using (var client = new WebClient())
            {
                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    if (!string.IsNullOrEmpty(domain))
                    {
                        client.Credentials = new NetworkCredential(username, password, domain);
                    }
                    else
                    {
                        client.Credentials = new NetworkCredential(username, password);
                    }
                }
                return client.DownloadData(sharePointUrl);
            }
        }

        public static MemoryStream Base64ToMemoryStream(string base64String)
        {
            // กรอง prefix เช่น "data:image/png;base64,"
            var base64Parts = base64String.Split(',');
            if (base64Parts.Length != 2)
            {
                throw new ArgumentException("Invalid base64 image format.");
            }

            // Decode base64 ไปเป็น byte[]
            byte[] imageBytes = Convert.FromBase64String(base64Parts[1]);

            // สร้าง MemoryStream
            return new MemoryStream(imageBytes);
        }

        public static string GetImageExtensionFromBase64(string base64String)
        {
            if (string.IsNullOrWhiteSpace(base64String))
                throw new ArgumentException("Base64 string is null or empty.");

            // ตัวอย่าง: data:image/png;base64,...
            if (base64String.StartsWith("data:"))
            {
                try
                {
                    var start = base64String.IndexOf("/") + 1;
                    var end = base64String.IndexOf(";", start);
                    string mime = base64String.Substring(start, end - start);

                    switch (mime.ToLower())
                    {
                        case "jpeg":
                        case "jpg":
                            return ".jpg";
                        case "png":
                            return ".png";
                        case "gif":
                            return ".gif";
                        case "bmp":
                            return ".bmp";
                        case "webp":
                            return ".webp";
                        default:
                            return ".bin"; // fallback
                    }
                }
                catch
                {
                    return ".bin"; // fallback ถ้ามี error
                }
            }

            return ".bin"; // ไม่ใช่ base64 แบบมี mime
        }
    }
}
