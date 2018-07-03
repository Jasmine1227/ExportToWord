using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportToWord.Common
{
    public class Result
    {
        /// <summary>
        /// 返回码
        /// </summary>
        public int code { get; set; }
        /// <summary>
        /// 返回信息
        /// </summary>
        public string message { get; set; }
        /// <summary>
        /// 图片地址
        /// </summary>
        public string imgUrl { get; set; }
    }
}