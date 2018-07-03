using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExportToWord.Common;
using System.IO;

namespace ExportToWord.Areas.ReportForms.Controllers
{
    public class ReportFormsController : Controller
    {
        //定义全局变量，存储位置
        string chartsPath = "/ChartsImg/";
        // GET: ReportForms/ReportForms
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 保存图片
        /// </summary>
        /// <param name="base64Info"></param>
        /// <param name="fileType">保存的图片类型</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Export(string base64Info, string fileType)
        {
            Result result = new Result();
            try
            {
                string[] arr = base64Info.Split(new string[] { "base64," }, StringSplitOptions.RemoveEmptyEntries);
                byte[] byteArray = Convert.FromBase64String(arr[1]);
                string path = AppDomain.CurrentDomain.BaseDirectory + chartsPath;
                //string path = Server.MapPath("/Resoucrces/File/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                //确保图片名称的唯一性
                Guid chartsID = Guid.NewGuid();
                //string filename = DateTime.Now.ToFileTime() + "." + fileType;
                string filename = chartsID + "." + fileType;
                string savePath = path + filename;

                FileStream fs = System.IO.File.Create(savePath);
                fs.Write(byteArray, 0, byteArray.Length);
                fs.Close();

                //string pdfPath = path + DateTime.Now.ToFileTime() + ".png";
                result.code = 1;
                result.message = "保存图片成功";
                //返回相对地址
                //_rsp.Data = FileTools._ReportChartsPath + filename;
                //返回绝对地址
                result.imgUrl = savePath;
            }
            catch (Exception ex)
            {
                result.code = -1;
                result.message = "引发异常";
            }


            //ConvertHelper.ConvertJPGToPDF(savePath, pdfPath);

            return Json(result, JsonRequestBehavior.AllowGet);
        }
    }
}