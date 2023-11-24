using Acrobat;
using OfficeAndPdfConverter.Controllers;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System;
using Microsoft.Office.Interop.Word;
namespace OfficeAndPdfConverter.Services
{
    public class AcrobatService
    {

        public static void PdfToAnyThing(string inputPDFPath, string outputWordPath, string fileType = FileFormat.Docx)
        {
            AcroPDDoc pdfd = new AcroPDDoc();
            pdfd.Open(inputPDFPath);
            Object jsObj = pdfd.GetJSObject();
            Type jsType = pdfd.GetType();
            object[] saveAsParam = { outputWordPath, fileType, "", false, false };
            var vrc = jsType.InvokeMember("saveAs",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, jsObj, saveAsParam,
                CultureInfo.InvariantCulture);
            pdfd.Close();
            Marshal.ReleaseComObject(pdfd);
        }

    }


    public class FileFormat
    {
        public const string Docx = "com.adobe.acrobat.docx";
        public const string Doc = "com.adobe.acrobat.doc";
        public const string Png = "com.adobe.acrobat.png";
        public const string Xlsx = "com.adobe.acrobat.xlsx";
        public const string Xls = "com.adobe.acrobat.spreadsheet";
        public const string Txt = "com.adobe.acrobat.accesstext";
        public const string Xml = "com.adobe.acrobat.xml-1-00";
        public const string Tiff = "com.adobe.acrobat.tiff";
        public const string Rft = "com.adobe.acrobat.rft";
        public const string PS = "com.adobe.acrobat.ps";
        public const string Jp2k = "com.adobe.acrobat.jp2k";
        public const string Jpeg = "com.adobe.acrobat.jpeg";
        public const string Html = "com.adobe.acrobat.html";
        public const string Eps = "com.adobe.acrobat.eps";
        public const string Pptx = "com.adobe.acrobat.pptx";
    }
}
