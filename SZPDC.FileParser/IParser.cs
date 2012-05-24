using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZPDC.FileParser
{
    /// <summary>
    /// 文档转换接口
    /// </summary>
    public interface IParser
    {
        void Parse(object src, object dest);
    }
}
