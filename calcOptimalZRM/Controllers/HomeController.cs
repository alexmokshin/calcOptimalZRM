using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;
using calcOptimalZRM.Models;
using System.Linq;
using System.Data.Entity;
using System.Net.Mime;
using Microsoft.Ajax.Utilities;
using WebGrease.Configuration;


namespace calcOptimalZRM.Controllers
{
    
    public class HomeController : Controller
    {
        private static void killDOZER(string name) //функция для убийства процесса
        {
            System.Diagnostics.Process[] etc = System.Diagnostics.Process.GetProcesses();//получим процессы
            foreach (System.Diagnostics.Process anti in etc)//обойдем каждый процесс
            {
                if (anti.ProcessName.ToLower().Contains(name.ToLower()))
                    anti.Kill();
            }
        }
        public static DateTime FirstDateTime;
        public static byte NomerPech;
        //объявим контекст подключения к базе данных повыше
        testReportEntities test = new testReportEntities();
        
        public ActionResult Index()
        {
            string dbConnTest =
                "data source=ASUSYASHA;initial catalog=Report;persist security info=True;user id=test_alexmoksh;password=654321;MultipleActiveResultSets=True;App=EntityFramework";
            test.Database.Connection.ConnectionString = dbConnTest;
            return View();
            
        }

        public ActionResult About()
        {
            ViewBag.Message = "Данное Web-приложение выполнено в рамках ВКР. Бакалавриат. Кафедра ТИМ";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        public ActionResult DbModel()
        {
            List<OptShihtDC_Shihta_Load_Result> opShLoad = new List<OptShihtDC_Shihta_Load_Result>();
            //testReportEntities test = new testReportEntities();
            //DateTime pickDateTime = DateTime.Parse("01." + "12" + "." + "2014");

            DateTime pechDateTime = FirstDateTime;
            byte pech = NomerPech;
            //var shihtaLoad = test.OptShihtDC_Shihta_Load(Convert.ToDateTime("2014-12-01"), 1, 0, 0);
            var shihtaLoad = test.OptShihtDC_Shihta_Load(pechDateTime, pech, 0, 0);
            //OptShihtDC_Param_Load_Result opths = new OptShihtDC_Param_Load_Result();
           // OptShihtDC_Shihta_Load_Result opths = new OptShihtDC_Shihta_Load_Result();
            
            
            foreach (var p in shihtaLoad)
            {
                OptShihtDC_Shihta_Load_Result opths = new OptShihtDC_Shihta_Load_Result();
                opths.dtFirstDay = p.dtFirstDay;
                opths.Материал = p.Материал;
                opths.Al2O3___ = p.Al2O3___;
                opths.CaO___ = p.CaO___;
                opths.Cr___ = p.Cr___;
                opths.Fe2O3___ = p.Fe2O3___;
                opths.FeO___ = p.FeO___;
                opths.Fe___ = p.Fe___;
                opths.H2O___ = p.H2O___;
                opths.MgO___ = p.MgO___;
                opths.MnO___ = p.MnO___;
                opths.PPP___ = p.PPP___;
                opths.P___ = p.P___;
                opths.S___ = p.S___;
                opths.SiO2___ = p.SiO2___;
                opths.TiO2___ = p.TiO2___;
                opths.ZnO___ = p.ZnO___;
                opths.Доля = p.Доля;
                opths.Код_материала = p.Код_материала;
                opths.Номер_печи = p.Номер_печи;
                opths.Расход__кг_т_чугуна = p.Расход__кг_т_чугуна;
                opths.Тип_материала = p.Тип_материала;
                opShLoad.Add(opths);
            }
            if (opShLoad.Count == 0)
            {
                return View("Error");
            }

            ViewBag.shload = opShLoad;
            return View();
        }
       

    
   
    public ActionResult PickData()
        {
            List< SelectListItem> actPech = new List<SelectListItem>();
            //testReportEntities test = new testReportEntities();

           // string dbConnTest =
           //     "data source=ASUSYASHA;initial catalog=Report;persist security info=True;user id=alexmoksh;password=654321;MultipleActiveResultSets=True;App=EntityFramework";
            bool ifExists = test.Database.Exists();
            string dbConn = test.Database.Connection.ConnectionString;
            
            test.Database.Connection.Open();
            //test.Database.Connection.
            var query = from dnp in test.DC_NSI_Pech where dnp.actual == true select dnp;
            foreach (var result in query)
            {
                actPech.Add(new SelectListItem {Text = result.Name, Value = Convert.ToString(result.PechId)});
            }

            ViewBag.PechType = actPech;
            List<SelectListItem> months = new List<SelectListItem>
            {
                new SelectListItem {Text = "Январь", Value = "01"},
                new SelectListItem {Text = "Февраль", Value = "02"},
                new SelectListItem {Text = "Март", Value = "03"},
                new SelectListItem {Text = "Апрель", Value = "04"},
                new SelectListItem {Text = "Май", Value = "05"},
                new SelectListItem {Text = "Июнь", Value = "06"},
                new SelectListItem {Text = "Июль", Value = "07"},
                new SelectListItem {Text = "Август", Value = "08"},
                new SelectListItem {Text = "Сентябрь", Value = "09"},
                new SelectListItem {Text = "Октябрь", Value = "10"},
                new SelectListItem {Text = "Ноябрь", Value = "11"},
                new SelectListItem {Text = "Декабрь", Value = "12"}
            };
            ViewBag.Monthes = months;
            test.Database.Connection.Close();
            return View();
        }
        [HttpPost]
        public ActionResult CalcZHRM(int Monthes, int PechType, string YearPick)
        {
            
            if ( YearPick.IsNullOrWhiteSpace())
            {
                return View("Error");
            }

            string excelPath = Server.MapPath("~/Content/" + "Оптимальная доменная шихта 2010_2.xlsm");
            DateTime pickDateTime = DateTime.Parse("01."+Convert.ToString(Monthes)+"."+YearPick);
            int npech = PechType;
            FirstDateTime = pickDateTime;
            NomerPech = Convert.ToByte(npech);
            List<OptShihtDC_Param_Load_Result> opShihtParam = new List<OptShihtDC_Param_Load_Result>();
            List<PechParams> paramsPechList = new List<PechParams>();
            //testReportEntities test = new testReportEntities();
            test.Database.Connection.Open();
            var shihtaParam = test.OptShihtDC_Param_Load(pickDateTime, 0);
            foreach (var p in shihtaParam)
            {
                PechParams pp = new PechParams();
                if (PechType == 1)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP1;
                }
                if (PechType == 2)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP2;
                }
                if (PechType == 3)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP3;
                }
                if (PechType == 4)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP4;
                }
                if (PechType == 5)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP5;
                }
                if (PechType == 6)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP6;
                }
                if (PechType == 7)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP7;
                }
                if (PechType == 8)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP8;
                }
                if (PechType == 9)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP9;
                }
                if (PechType == 10)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP10;
                }
                paramsPechList.Add(pp);
            }

            ViewBag.ParamsPech = paramsPechList;
            OptZRM opz = new OptZRM();
            opz.nachDataPech = pickDateTime;
            opz.nomerPech = Convert.ToByte(npech);
            opz.SetDataOnPechInExcel(excelPath);
            DomCehParameter domCehParameter = new DomCehParameter();
            domCehParameter.proizvDomPechModel = Convert.ToDouble(opz.proizvDomPechModel);
            domCehParameter.summUdRashKoksModel = Convert.ToDouble(opz.summUdRashKoksModel);
            domCehParameter.sodSeraChugunModel = Convert.ToDouble(opz.sodSeraChugunModel);
            domCehParameter.dolAglomeratMmkModel = Convert.ToDouble(opz.dolAglomeratMmkModel);
            domCehParameter.Al2O3ShlakModel = Convert.ToDouble(opz.Al2O3ShlakModel);
            domCehParameter.MgOShlakModel = Convert.ToDouble(opz.MgOShlakModel);
            ViewBag.DomCehPar = domCehParameter;
            VvodOgranicheniy vvod = new VvodOgranicheniy();
            vvod.aglmmkmax = opz.aglmmkmax;
            vvod.aglmmkmin = opz.aglmmkmin;
            vvod.aglmmknow = opz.aglmmknow;
            
            vvod.grdvzkkkonshlak = opz.grdvzkkkonshlak;
            vvod.grdvzkkkonshlakmax = opz.grdvzkkkonshlakmax;
            vvod.grdvzkkkonshlakmin = opz.grdvzkkkonshlakmin;
            
            vvod.onskonshlakkzad = opz.onskonshlakkzad;
            vvod.onskonshlakmin = opz.onskonshlakmin;
            vvod.osnkonshlack = opz.osnkonshlack;
            vvod.osnkonshlakmax = opz.osnkonshlakmax;
            vvod.summdoleyzrmnow = opz.summdoleyzrmnow;
            vvod.vzskkonshlak = opz.vzskkonshlak;
            vvod.vzskkonshlak1450 = opz.vzskkonshlak1450;
            vvod.vzskkonshlakmax = opz.vzskkonshlakmax;
            vvod.vzskkonshlakmax1450 = opz.vzskkonshlakmax1450;
            vvod.vzskkonshlakmin = opz.vzskkonshlakmin;
            vvod.vzskkonshlakmin1450 = opz.vzskkonshlakmin1450;
           
            ViewBag.Vvod = vvod;
            test.Database.Connection.Close();
            return View();

        }

        [HttpPost]
        public ActionResult ResultCalcZRM(VvodOgranicheniy vvod, FormCollection frm)
        {
            int targetType;
            int influence;
            var d = vvod.vzskkonshlakmax;
            try
            {
                targetType = Convert.ToInt32(frm["target"].ToString());
                influence = Convert.ToInt32(frm["vozdeistv"].ToString());
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.ToString();
                return View("Error");
            }
            string[] conclusion =
                {
                    frm["first_concl"], frm["second_concl"], frm["third_concl"], frm["fourth_concl"], frm["fifth_concl"]
                };
            bool[] b = new bool[conclusion.Length];
            for (int i = 0; i < conclusion.Length; i++)
            {
                if (conclusion[i].Contains("true"))
                {
                    b[i] = true;
                }
                else
                {
                    b[i] = false;
                } 

            }
            string excelPath = Server.MapPath("~/Content/" + "Оптимальная доменная шихта 2010_2.xlsm");

            //string excelPath = Server.MapPath("~/Content/" + "Оптимальная доменная шихта 2010_.xlt");

            OptZRM opz = new OptZRM();
            opz.CalcOptZRM(b,targetType,influence,excelPath);
            Okatishi okatishi = new Okatishi();

            okatishi.okatSsgokModel = opz.okatSsgok;
            okatishi.okatLebedinskModel = opz.okatLebedinsk;
            okatishi.okatKachanarskieModel = opz.okatKachanarskie;
            okatishi.okatMihailovskModel = opz.okatMihailovsk;
            vvod.summdoleyzrmnow = opz.summdoleyzrmnow;
            vvod.vzskkonshlak = opz.vzskkonshlak;
            vvod.vzskkonshlak1450 = opz.vzskkonshlak1450;
            vvod.aglmmknow = opz.aglmmknow;
            vvod.osnkonshlack = opz.osnkonshlack;
            ViewBag.Vihod = vvod;
            ViewBag.Okatish = okatishi;
            killDOZER("excel");
            return View();
        }


    }
}