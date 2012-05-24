using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZPDC.FileParser
{
    using Microsoft.Office.Interop.PowerPoint;

    public class PPTParser : IParser
    {
        private string _document;

        Microsoft.Office.Interop.PowerPoint.Application pptApp = 
			new Microsoft.Office.Interop.PowerPoint.Application();

        Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
        object missing = Type.Missing;
        /// <summary>
        /// word文档路径
        /// </summary>
        public string FilePath
        {
            get { return _document; }
            set { _document = value; }
        }

        public PPTParser()
        {

        }
        public PPTParser(string path)
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
                presentation = pptApp.Presentations.Open(src.ToString(), Microsoft.Office.Core.MsoTriState.msoTrue
                , Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);


                presentation.ExportAsFixedFormat(dest.ToString(), PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentScreen
                    , Microsoft.Office.Core.MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides
                    , Microsoft.Office.Core.MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, string.Empty, false
                    , false, false, true, true, missing);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    presentation = null;
                }

                if (pptApp != null)
                {
                    pptApp.Quit();
                    pptApp = null;
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
