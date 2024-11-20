using ceTe.DynamicPDF.Conversion;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using System.Reflection.Metadata;
using Microsoft.AspNetCore.Hosting;
using Aspose.Words.LowCode;
using GrapeCity.Documents.Word;
using GrapeCity.Documents.Word.Layout;
using System.IO.Compression;
namespace uploadAPI.Controllers
{
    [EnableCors("AllowSpecificOrigin")]
    [ApiController]
    [Route("[controller]")]
    public class UploadConvertController : ControllerBase
    {

        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment _hostingEnvironment;

        public UploadConvertController(Microsoft.AspNetCore.Hosting.IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
       

        [HttpPost("convertToPdf")]
        public IActionResult ConvertToPdf(IFormFile file)
        {
            string rootPath = _hostingEnvironment.ContentRootPath;
            string uploadFolderPath = Path.Combine(rootPath,  "files");

            string uploadFolderPathConvert = Path.Combine(rootPath,"files_convert");
            try
            {
                if (!Directory.Exists(uploadFolderPath))
                {
                    Directory.CreateDirectory(uploadFolderPath);
                }
                if (!Directory.Exists(uploadFolderPathConvert))
                {
                    Directory.CreateDirectory(uploadFolderPathConvert);
                }
                if (file == null || file.Length == 0)
                {
                    return BadRequest("File không hợp lệ hoặc bị rỗng.");
                }
                string filePath = Path.Combine(uploadFolderPath, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                Console.WriteLine($"File input: {filePath}");
                string nameFile = Path.GetFileName(filePath);

                string nameConvert = nameFile.Substring(0, nameFile.IndexOf('.'));

                string fileNameConvert = $"{uploadFolderPathConvert}\\{nameConvert}.pdf";

                var wordDoc = new GcWordDocument();

                wordDoc.Load(filePath);

                using (var layout = new GcWordLayout(wordDoc))
                {
                    PdfOutputSettings pdfOutputSettings = new PdfOutputSettings();
                    pdfOutputSettings.CompressionLevel = CompressionLevel.Fastest;
                    pdfOutputSettings.ConformanceLevel = GrapeCity.Documents.Pdf.PdfAConformanceLevel.PdfA1a;
                    layout.SaveAsPdf(fileNameConvert, null, pdfOutputSettings);
                }
                //WordConverter converter = new WordConverter($"{uploadFolderPath}\\{nameFile}");
                //converter.Convert(fileNameConvert);

                byte[] fileBytes = System.IO.File.ReadAllBytes(fileNameConvert);
                string name = file.FileName.Substring(0, nameFile.IndexOf('.'));
                System.IO.File.Delete(filePath);
                return File(fileBytes, "application/pdf", $"{name}.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return BadRequest($"Conversion failed: {ex.Message}");
            }
            finally
            {
                string[] files = Directory.GetFiles(uploadFolderPathConvert);
                foreach (var fileconvert in files)
                {
                    System.IO.File.Delete(fileconvert);
                }
            }
            
            }
        //[HttpPost("convertToPdfv1")]
        //public IActionResult ConvertToPdfV1(IFormFile file)
        //{
        //    string rootPath = _hostingEnvironment.ContentRootPath;
        //    string uploadFolderPath = Path.Combine(rootPath, "files");
            
        //    string uploadFolderPathConvert = Path.Combine(rootPath, "files_convert");
        //    try
        //    {
        //        if (!Directory.Exists(uploadFolderPath))
        //        {
        //            Directory.CreateDirectory(uploadFolderPath);
        //        }
        //        if (!Directory.Exists(uploadFolderPathConvert))
        //        {
        //            Directory.CreateDirectory(uploadFolderPathConvert);
        //        }
        //        if (file == null || file.Length == 0)
        //        {
        //            return BadRequest("File không hợp lệ hoặc bị rỗng.");
        //        }
        //        string filePath = Path.Combine(uploadFolderPath, file.FileName);

        //        using (var stream = new FileStream(filePath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //        }
        //        Console.WriteLine($"File input: {filePath}");
        //        string nameFile = Path.GetFileName(filePath);

        //        string nameConvert = nameFile.Substring(0, nameFile.IndexOf('.'));

        //        string fileNameConvert = $"{uploadFolderPathConvert}\\{nameConvert}.pdf";

        //        var wordDoc = new GcWordDocument();

        //        wordDoc.Load(filePath);

        //        using(var layout = new GcWordLayout(wordDoc))
        //        {
        //            PdfOutputSettings pdfOutputSettings = new PdfOutputSettings();
        //            pdfOutputSettings.CompressionLevel = CompressionLevel.Fastest;
        //            pdfOutputSettings.ConformanceLevel = GrapeCity.Documents.Pdf.PdfAConformanceLevel.PdfA1a;
        //            layout.SaveAsPdf(fileNameConvert, null, pdfOutputSettings);
        //        }
        //        //WordConverter converter = new WordConverter($"{uploadFolderPath}\\{nameFile}");
        //        //converter.Convert(fileNameConvert);

        //        byte[] fileBytes = System.IO.File.ReadAllBytes(fileNameConvert);
        //        string name = file.FileName.Substring(0, nameFile.IndexOf('.'));
        //        System.IO.File.Delete(filePath);
        //        return File(fileBytes, "application/pdf", $"{name}.pdf");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex);
        //        return BadRequest($"Conversion failed: {ex.Message}");
        //    }
        //    finally
        //    {
        //        string[] files = Directory.GetFiles(uploadFolderPathConvert);
        //        foreach (var fileconvert in files)
        //        {
        //            System.IO.File.Delete(fileconvert);
        //        }
        //    }

        //}
        [HttpPost("convertToPdf1")]
        public IActionResult ConvertToPdf1(FileInfo _form)
        {
            string rootPath = _hostingEnvironment.ContentRootPath;
            string uploadFolderPath = Path.Combine(rootPath, "files");
            string uploadFolderPathConvert = Path.Combine(rootPath, "files_convert");
            try
            {
                //if (!Directory.Exists(Server.MapPath("~/Assets/" + _form.folder_name)))
                //{
                //    Directory.CreateDirectory(Server.MapPath("~/Assets/" + _form.folder_name));
                //}
                if (!Directory.Exists(uploadFolderPath))
                {
                    Directory.CreateDirectory(uploadFolderPath);
                }
                if (!Directory.Exists(uploadFolderPathConvert))
                {
                    Directory.CreateDirectory(uploadFolderPathConvert);
                }
                if (_form.filecontent == null || _form.filecontent.Length == 0)
                {
                    return BadRequest("File không hợp lệ hoặc bị rỗng.");
                }
                string filePath = Path.Combine(uploadFolderPath, _form.file_name);

                System.IO.File.WriteAllBytes(filePath, Convert.FromBase64String(_form.filecontent));

                string nameFile = Path.GetFileName(filePath);

                string nameConvert = nameFile.Substring(0, nameFile.IndexOf('.'));
                //string nameConvert = Path.GetFileNameWithoutExtension(nameFile);
                string fileNameConvert = $"{uploadFolderPathConvert}\\{nameConvert}.pdf";

                var wordDoc = new GcWordDocument();

                wordDoc.Load(filePath);

                using (var layout = new GcWordLayout(wordDoc))
                {
                    PdfOutputSettings pdfOutputSettings = new PdfOutputSettings();
                    pdfOutputSettings.CompressionLevel = CompressionLevel.Fastest;
                    pdfOutputSettings.ConformanceLevel = GrapeCity.Documents.Pdf.PdfAConformanceLevel.PdfA1a;
                    layout.SaveAsPdf(fileNameConvert, null, pdfOutputSettings);
                }
                //WordConverter converter = new WordConverter($"{uploadFolderPath}\\{nameFile}");
                //converter.Convert(fileNameConvert);

                byte[] pdfData = System.IO.File.ReadAllBytes(fileNameConvert);

                string base64Pdf = Convert.ToBase64String(pdfData);
                System.IO.File.Delete(filePath);

                return Content(base64Pdf, "application/json");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return BadRequest($"Conversion failed: {ex.Message}");
            }
            finally
            {
                string[] files = Directory.GetFiles(uploadFolderPathConvert);
                foreach (var fileconvert in files)
                {
                    System.IO.File.Delete(fileconvert);
                }
            }

        }
    }
    public class FileInfo
    {
        //public string? folder_name { get; set; }
        public string? file_name { get; set; }
        public string? filecontent { get; set; }
    }
}

            
            
