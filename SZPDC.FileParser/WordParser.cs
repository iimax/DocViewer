using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZPDC.FileParser
{
    using Microsoft.Office.Interop.Word;

    public class WordParser : IParser
    {
        private string _document;

        private ApplicationClass wordApp = new ApplicationClass();

        private Document doc = null;
        object paraMissing = Type.Missing;
        /// <summary>
        /// word文档路径
        /// </summary>
        public string FilePath
        {
            get { return _document; }
            set { _document = value; }
        }

        public WordParser()
        {

        }

        public WordParser(string path)
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
                doc = wordApp.Documents.Open(ref src, ref paraMissing, ref paraMissing
                    , ref paraMissing, ref paraMissing
                    , ref paraMissing, ref paraMissing
                    , ref paraMissing, ref paraMissing, ref paraMissing, ref paraMissing
                    , ref paraMissing, ref paraMissing, ref paraMissing, ref paraMissing
                    , ref paraMissing);

                
                int startPage = 0;
                int endPage = doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref paraMissing);

                if (doc != null)
                {
                    doc.ExportAsFixedFormat(dest.ToString(), WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen
                        , WdExportRange.wdExportAllDocument, startPage, endPage, WdExportItem.wdExportDocumentWithMarkup,
                        true, true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, true, ref paraMissing);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(ref paraMissing, ref paraMissing, ref paraMissing);
                    doc = null;
                }

                if (wordApp != null)
                {
                    wordApp.Quit(ref paraMissing, ref paraMissing, ref paraMissing);
                    wordApp = null;
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
