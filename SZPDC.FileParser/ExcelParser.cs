using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZPDC.FileParser
{
    using Microsoft.Office.Interop.Excel;
    public class ExcelParser : IParser
    {
        private string _document;

        Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new ApplicationClass();
        Microsoft.Office.Interop.Excel.Workbook xlsx = null;
        object missing = Type.Missing;
        /// <summary>
        /// word文档路径
        /// </summary>
        public string FilePath
        {
            get { return _document; }
            set { _document = value; }
        }

        public ExcelParser()
        {

        }

        public ExcelParser(string path)
        {
            this.FilePath = path;
        }
        #region IParser 成员

        public void Parse(object src, object dest)
        {
            if (src == null || string.IsNullOrEmpty(src.ToString()))
            {
                throw new ArgumentNullException("源文件不能为空");
            }
            if (dest == null || string.IsNullOrEmpty(dest.ToString()))
            {
                throw new ArgumentNullException("目标路径不能为空");
            }

            try
            {
                xlsx = excelApp.Workbooks.Open(src.ToString(), missing, missing, missing, missing, missing
                    , missing, missing, missing, missing, missing, missing, missing, missing, missing);

                
                int startPage = 0;
                int endPage = 10;
                if (xlsx != null)
                {
                    xlsx.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, dest, missing, missing, missing, missing, missing, missing,
                        missing);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (xlsx != null)
                {
                    xlsx.Close(false, missing, missing);
                    xlsx = null;
                }

                if (excelApp!= null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion
    }
}
