using OfficeAndPdfConverter.Models;
using OfficeAndPdfConverter.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace OfficeAndPdfConverter.Controllers
{
    public class FileController : Controller
    {
        IWebHostEnvironment _hostEnvironment = null;

        public FileController(IWebHostEnvironment hostEnvironment)
        {
            _hostEnvironment = hostEnvironment;
        }

        [HttpGet]
        public IActionResult Index(string fileName = "")
        {
            FilesModel fileObj = new FilesModel();
            fileObj.Name = fileName;

            string path = $"{_hostEnvironment.WebRootPath}\\files\\";
            int nId = 1;
            ScanFolder(fileObj, path, "*.pdf", 1, ref nId);
            ScanFolder(fileObj, path, "*.docx", 2, ref nId);
            ScanFolder(fileObj, path, "*.doc", 2, ref nId);
            ScanFolder(fileObj, path, "*.pptx", 3, ref nId);
            ScanFolder(fileObj, path, "*.ppt", 3, ref nId);
            ScanFolder(fileObj, path, "*.xlsx", 4, ref nId);
            ScanFolder(fileObj, path, "*.xls", 4, ref nId);
            return View(fileObj);
        }

        private static int ScanFolder(FilesModel fileObj, string path, string type, int typeInt, ref int nId)
        {
            foreach (string pdfPath in Directory.EnumerateFiles(path, type))
            {
                fileObj.Files.Add(new FilesModel()
                {
                    FileId = nId++,
                    Name = Path.GetFileName(pdfPath),
                    Path = pdfPath,
                    Type = typeInt
                });
            }
            return nId;
        }

        [HttpPost]
        public IActionResult Index(IFormFileCollection files, [FromServices] IWebHostEnvironment hostEnvironment)
        {
            FilesModel fileObj = new FilesModel();
            if (files?.Count > 0)
            {
                foreach (var file in files)
                {
                    string fileName = $"{hostEnvironment.WebRootPath}\\files\\{file.FileName}";
                    using (FileStream fileStream = System.IO.File.Create(fileName))
                    {
                        file.CopyTo(fileStream);
                        fileStream.Flush();
                    }
                }
            }
            else
            {
                fileObj.Inofmation = "请选中一个文件";
            }
            return Index();
        }

  

        public IActionResult ToDocx(string fileName)
        {
            return ToAnyThing(fileName, ".docx", FileFormat.Docx);
        }

        public IActionResult ToDoc(string fileName)
        {
            return ToAnyThing(fileName, ".doc", FileFormat.Doc);
        }


        public IActionResult ToHtml(string fileName)
        {
            return ToAnyThing(fileName,".html", FileFormat.Html);
        }

        public IActionResult ToJpg(string fileName)
        {
            return ToAnyThing(fileName, ".jpg", FileFormat.Jpeg);
        }

        public IActionResult ToXlsx(string fileName)
        {
            return ToAnyThing(fileName, ".xlsx", FileFormat.Xlsx);
        }

        public IActionResult ToPptx(string fileName)
        {
            return ToAnyThing(fileName, ".pptx", FileFormat.Pptx);
        }

        public IActionResult SrcFile(string fileName)
        {
            string path = _hostEnvironment.WebRootPath + "\\files\\" + fileName;
            return File(System.IO.File.ReadAllBytes(path), "application/octet-stream", fileName);
        }

        private IActionResult ToAnyThing(string fileName, string fileExtension, string fileType)
        {
            string path = _hostEnvironment.WebRootPath + "\\files\\" + fileName;
            string nameOnly = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            var newName = nameOnly + fileExtension;
            var newPath = _hostEnvironment.WebRootPath + "\\newfiles\\" + newName;
            AcrobatService.PdfToAnyThing(path, newPath, fileType);
            return File(System.IO.File.ReadAllBytes(newPath), "application/octet-stream", newName);
        }

        public IActionResult ToPDF(string fileName, int type)
        {
            string path = _hostEnvironment.WebRootPath + "\\files\\" + fileName;
            string nameOnly = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            var newName = nameOnly + ".pdf";
            var newPath = _hostEnvironment.WebRootPath + "\\newfiles\\" + newName;

            if (type == 2)
            {
                OfficeService.Word2Pdf(path, newPath);
            }
            else if (type == 3)
            {
                OfficeService.Ppt2Pdf(path, newPath);
            }
            else if (type == 4)
            {
                OfficeService.Excel2Pdf(path, newPath);
            }

            return File(System.IO.File.ReadAllBytes(newPath), "application/octet-stream", newName);
        }
    }




}
