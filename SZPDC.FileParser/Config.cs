using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZPDC.FileParser
{
    /// <summary>
    /// 配置表
    /// </summary>
    public class Config
    {
        /// <summary>
        /// pdf2swf.exe路径
        /// </summary>
        public static string PDF2SWF_PATH = System.Configuration.ConfigurationManager.AppSettings["PDF2SWF"];
    }
}
