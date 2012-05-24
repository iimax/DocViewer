using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace SZPDC.FileParser
{
    public class PDFParser : IParser
    {
        private string _document;
        /// <summary>
        /// PDF文档路径
        /// </summary>
        public string FilePath
        {
            get { return _document; }
            set { _document = value; }
        }

        #region IParser 成员

        public void Parse(object src, object dest)
        {
            if (src == null || string.IsNullOrEmpty(src.ToString()))
            {
                throw new ArgumentNullException("源文件不能为空！");
            }
            if (dest == null || string.IsNullOrEmpty(dest.ToString()))
            {
                throw new ArgumentNullException("目标路径不能为空！");
            }
            if (!File.Exists(src.ToString()))
            {
                throw new ArgumentException("源文件不存在！");
            }

            if (File.Exists(dest.ToString()))
            {
                throw new ArgumentException(string.Format("目标文件 {0} 已存在！", dest.ToString()));
            }
            try
            {
                //将pdf文档转成temp.swf文件
                string cmd = String.Format("\"{0}\" -o \"{1}\" -t -s flashversion=9"
                    //,Config.PDF2SWF_PATH
                     , src.ToString()
                     , dest.ToString());
                RunShell(cmd);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
            }
        }

        #endregion

        /// <summary>
        /// 运行命令
        /// </summary>
        /// <param name="strShellCommand">命令字符串</param>
        private static void RunShell(string strShellCommand)
        {
            Process cmd = Process.Start(Config.PDF2SWF_PATH);
            cmd.StartInfo.Arguments = strShellCommand;
            cmd.Start();
        }
    }
}
