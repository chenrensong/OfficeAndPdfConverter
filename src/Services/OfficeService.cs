using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
namespace OfficeAndPdfConverter.Services
{
    public class OfficeService
    {

        public static bool Word2Pdf(string sourcePath, string savePath)
        {
            bool result = false;
            Word.Document word = null;
            Word.Application wordApplication = null;
            object missing = Type.Missing;

            try
            {
                if (wordApplication == null)
                {
                    wordApplication = new Word.Application();
                }
                word = wordApplication.Documents.Open(sourcePath);
                word.ExportAsFixedFormat(savePath, Word.WdExportFormat.wdExportFormatPDF);
                result = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (word != null)
                {
                    word.Close();
                    word = null;
                }

                if (wordApplication != null)
                {
                    wordApplication.Quit();
                    wordApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }
        public static bool Excel2Pdf(string sourcePath, string savePath)
        {
            bool result = false;
            Excel.Workbook presentation = null;
            Excel.Application pptApplication = null;
            var targetFileType = Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;

            try
            {
                pptApplication = new Excel.Application();
                presentation = pptApplication.Workbooks.Open(sourcePath);
                presentation.ExportAsFixedFormat(targetFileType, savePath);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    presentation = null;
                }

                if (pptApplication != null)
                {
                    pptApplication.Quit();
                    pptApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }

        /// <summary>
        /// 转换函数
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="savePath">保存的路径</param>
        /// <returns></returns>
        public static bool Ppt2Pdf(string sourcePath, string savePath)
        {
            bool result = false;
            PowerPoint.Presentation presentation = null;
            PowerPoint.Application pptApplication = null;
            object missing = Type.Missing;

            try
            {
                pptApplication = new PowerPoint.Application();
                presentation = pptApplication.Presentations.Open(sourcePath);
                presentation.ExportAsFixedFormat(savePath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    presentation = null;
                }

                if (pptApplication != null)
                {
                    pptApplication.Quit();
                    pptApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }

    }
}
