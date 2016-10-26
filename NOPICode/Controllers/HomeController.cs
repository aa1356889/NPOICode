using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NOPICode.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 多列表头并且合并单元格导出
        /// </summary>
        /// <returns></returns>
        public ActionResult Down1()
        {
            List<TestExportData> datas = new List<TestExportData>() { 
             new TestExportData(){UserName="小明",Address="tests地址1",Age="12",bloodSugar="2333",HeartRate="332"},
             new TestExportData(){UserName="小明2",Address="tests地址12",Age="12",bloodSugar="2333",HeartRate="332"},
             new TestExportData(){UserName="小明3",Address="tests地址4",Age="12",bloodSugar="33",HeartRate="33"},
             new TestExportData(){UserName="小明4",Address="tests地址5",Age="12",bloodSugar="44",HeartRate="44"}
            };
            string[,] header = new string[2, 5];
            header[0, 0] = "用户基本信息";
            header[0, 3] = "用户身体状况";
            header[1, 0] = "用户名字";
            header[1, 1] = "用户年龄";
            header[1, 2] = "用户地址";
            header[1, 3] = "心率";
            header[1, 4] = "血糖";
            var bytes = NOPIHelper.Export<TestExportData>(header, datas, null);
            return File(bytes, "application/vnd.ms-excel", "用户信息.xls");
        }

        /// <summary>
        /// 多列表头 过滤部分数据
        /// </summary>
        /// <returns></returns>
        public ActionResult Down2()
        {
            List<TestExportData2> datas = new List<TestExportData2>() { 
             new TestExportData2(){UserName="小明",Address="tests地址1",Age="12",bloodSugar="2333",HeartRate="332",id=1,Sex=0},
             new TestExportData2(){UserName="小明2",Address="tests地址12",Age="12",bloodSugar="2333",HeartRate="332",id=2,Sex=0},
             new TestExportData2(){UserName="小明3",Address="tests地址4",Age="12",bloodSugar="33",HeartRate="33",id=3,Sex=1},
             new TestExportData2(){UserName="小明4",Address="tests地址5",Age="12",bloodSugar="44",HeartRate="44",id=4,Sex=1}
            };
            string[,] header = new string[2, 5];
            header[0, 0] = "用户基本信息";
            header[0, 3] = "用户身体状况";
            header[1, 0] = "用户名字";
            header[1, 1] = "用户年龄";
            header[1, 2] = "用户地址";
            header[1, 3] = "心率";
            header[1, 4] = "血糖";
            var bytes = NOPIHelper.Export<TestExportData2>(header, datas, c => new {c.UserName,c.Address,c.Age,c.bloodSugar,c.HeartRate });
            return File(bytes, "application/vnd.ms-excel", "用户信息.xls");
        }

        /// <summary>
        /// 普通导出
        /// </summary>
        /// <returns></returns>
        public ActionResult Down3()
        {
            List<TestExportData2> datas = new List<TestExportData2>() { 
             new TestExportData2(){UserName="小明",Address="tests地址1",Age="12",bloodSugar="2333",HeartRate="332",id=1,Sex=0},
             new TestExportData2(){UserName="小明2",Address="tests地址12",Age="12",bloodSugar="2333",HeartRate="332",id=2,Sex=0},
             new TestExportData2(){UserName="小明3",Address="tests地址4",Age="12",bloodSugar="33",HeartRate="33",id=3,Sex=1},
             new TestExportData2(){UserName="小明4",Address="tests地址5",Age="12",bloodSugar="44",HeartRate="44",id=4,Sex=1}
            };
            string[,] header = new string[1, 7];
            header[0, 0] = "用户名字";
            header[0, 1] = "用户年龄";
            header[0, 2] = "用户地址";
            header[0, 3] = "心率";
            header[0, 4] = "血糖";
            header[0, 5] = "性别";
            header[0, 6] = "id";
            var bytes = NOPIHelper.Export<TestExportData2>(header, datas, c =>new {c.UserName,c.Address,c.Age,c.bloodSugar,c.HeartRate,c.id,Sex=(c.Sex==1?"男":"女") });
            return File(bytes, "application/vnd.ms-excel", "用户信息.xls");
        }


        public ActionResult Down4()
        {
             List<TestExportData2> datas = new List<TestExportData2>() { 
             new TestExportData2(){UserName="小明",Address="tests地址1",Age="12",bloodSugar="2333",HeartRate="332",id=1,Sex=0},
             new TestExportData2(){UserName="小明2",Address="tests地址12",Age="12",bloodSugar="2333",HeartRate="332",id=2,Sex=0},
             new TestExportData2(){UserName="小明3",Address="tests地址4",Age="12",bloodSugar="33",HeartRate="33",id=3,Sex=1},
             new TestExportData2(){UserName="小明4",Address="tests地址5",Age="12",bloodSugar="44",HeartRate="44",id=4,Sex=1}
            };
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("address", "南方医院");
            dic.Add("phone", "13243868978");
            dic.Add("username", "李强");
            dic.Add("phoneb", "13128273410");
            dic.Add("user", "小明");
            dic.Add("exportdate", DateTime.Now.ToString());
            var bytes = NOPIHelper.Export<TestExportData2>(HttpContext.Server.MapPath("/Temp/测试单模板.xls"), 3, dic, datas, datas.Count,
                c => new {c.UserName, c.Address, c.Age, c.bloodSugar, c.HeartRate});
           return File(bytes, "application/vnd.ms-excel", "用户信息.xls");
        }
    }
}
