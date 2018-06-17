using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;
using calcOptimalZRM.Models;
using System.Linq;
using System.Data.Entity;
using System.IO;
using System.Net.Mime;
using Microsoft.Ajax.Utilities;
using WebGrease.Configuration;
using System.Threading;


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

            DateTime pickDateTime;
            string excelPath = Server.MapPath("~/Content/" + "Оптимальная доменная шихта 2010_2.xlsm");
            try
            {
                pickDateTime = DateTime.Parse("01." + Convert.ToString(Monthes) + "." + YearPick);
            }
            catch (Exception e)
            {
                ViewBag.Error = e.ToString();
                return View("Error");
            }
            
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

            okatishi.okatSsgokModel = Convert.ToDouble(opz.okatSsgok);
            ResOkatSsgok = okatishi.okatSsgokModel;
            okatishi.okatLebedinskModel = Convert.ToDouble(opz.okatLebedinsk);
            ResOkatLebedinsk = okatishi.okatLebedinskModel;
            okatishi.okatKachanarskieModel = Convert.ToDouble(opz.okatKachanarskie);
            ResOkatKachanarsk = okatishi.okatKachanarskieModel;
            okatishi.okatMihailovskModel = Convert.ToDouble(opz.okatMihailovsk);
            ResOkatMihailovsk = okatishi.okatMihailovskModel;
            ParamShlakAfterCalc psac = new ParamShlakAfterCalc();
            psac.summdoleyzrmnow = opz.summdoleyzrmnow;
            ResSummDoleyZrm = psac.summdoleyzrmnow;
            psac.vzskkonshlak = opz.vzskkonshlak;
            ResVyazkKonShlak = psac.vzskkonshlak;
            psac.vzskkonshlak1450 = opz.vzskkonshlak1450;
            ResVyazkKonShlak14 = psac.vzskkonshlak1450;
            psac.aglmmknow = opz.aglmmknow;
            ResAglomMmk = psac.aglmmknow;
            psac.osnkonshlack = opz.osnkonshlack;
            ResOsnovKonShlak = psac.osnkonshlack;
            psac.grdvzkkkonshlak = opz.grdvzkkkonshlak;
            ResGradVyazkShlak = psac.grdvzkkkonshlak;
            ViewBag.Result = psac;
            ViewBag.Vihod = vvod;
            ViewBag.Okatish = okatishi;
            killDOZER("excel");
            return View();
        }

        private static double _resOkatSsgok;
        private static double _resOkatLebedinsk;
        private static double _resOkatKachanarsk;
        private static double _resOkatMihailovsk;
        private static double? _resSummDoleyZrm;
        private static double? _resVyazkKonShlak;
        private static double? _resVyazkKonShlak14;
        private static double? _resAglomMmk;
        private static double? _resOsnovKonShlak;
        private static double? _resGradVyazkShlak;

        public static double ResOkatSsgok { get => _resOkatSsgok; set => _resOkatSsgok = value; }
        public static double ResOkatLebedinsk { get => _resOkatLebedinsk; set => _resOkatLebedinsk = value; }
        public static double ResOkatKachanarsk { get => _resOkatKachanarsk; set => _resOkatKachanarsk = value; }
        public static double ResOkatMihailovsk { get => _resOkatMihailovsk; set => _resOkatMihailovsk = value; }
        public static double? ResSummDoleyZrm { get => _resSummDoleyZrm; set => _resSummDoleyZrm = value; }
        public static double? ResVyazkKonShlak { get => _resVyazkKonShlak; set => _resVyazkKonShlak = value; }
        public static double? ResVyazkKonShlak14 { get => _resVyazkKonShlak14; set => _resVyazkKonShlak14 = value; }
        public static double? ResAglomMmk { get => _resAglomMmk; set => _resAglomMmk = value; }
        public static double? ResOsnovKonShlak { get => _resOsnovKonShlak; set => _resOsnovKonShlak = value; }
        public static double? ResGradVyazkShlak { get => _resGradVyazkShlak; set => _resGradVyazkShlak = value; }

        public FileResult ExportExcel(Okatishi oks, ParamShlakAfterCalc pscalc)
        {
            string newname;
            test.Database.Connection.Open();
            List<OptShihtDC_Param_Load_Result> opShihtParam = new List<OptShihtDC_Param_Load_Result>();
            List<PechParams> paramsPechList = new List<PechParams>();
            var shihtaParam = test.OptShihtDC_Param_Load(FirstDateTime, 0);
            byte PechType = NomerPech;
            const string template = "Report.xlsx";
            newname = "Report_" + DateTime.Now.ToShortDateString()+DateTime.Now.Hour+DateTime.Now.Minute+DateTime.Now.Second + ".xlsx";
            string path = Server.MapPath("~/Content/" + template);
            string newpath = Server.MapPath("~/Content/") + newname;
            System.IO.File.Copy(path,newpath,false);

            ExcelReport exc = new ExcelReport(newpath, false);
            exc.ChangeWorkSheet("Параметры плавки");
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

            string dataparameter = "ДП-" + PechType + " Дата " + FirstDateTime.Date;
            exc.Write(1,1,dataparameter);
            exc.Write(2, 1, "Параметры плавки");
            for (int i =0;i<paramsPechList.Count;i++)
            {
                int p = i + 3;
                exc.Write(p,1,paramsPechList[i].Descr);
                exc.Write(p,2,paramsPechList[i].Val);
            }
            exc.ChangeWorkSheet("Параметры шихты_шлака");
            exc.Write(1,1,dataparameter);
            List<OptShihtDC_HimZolaKoks_Load1_Result> himzolaList = new List<OptShihtDC_HimZolaKoks_Load1_Result>();
            var himzolakok = test.OptShihtDC_HimZolaKoks_Load1(FirstDateTime, 0);
            var shihtaLoad = test.OptShihtDC_Shihta_Load(FirstDateTime, NomerPech, 0, 0);
            List<OptShihtDC_Shihta_Load_Result> opShLoad = new List<OptShihtDC_Shihta_Load_Result>();
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
            exc.Write(2,1,"Химический состав шихты");
            for (int i = 0; i < opShLoad.Count; i++)
            {
                int p = i + 4;
                exc.Write(p,1,opShLoad[i].Материал);
                exc.Write(p, 2, opShLoad[i].Расход__кг_т_чугуна);
                exc.Write(p, 3, opShLoad[i].Доля);
                exc.Write(p, 4, opShLoad[i].Fe___);
                exc.Write(p, 5, opShLoad[i].FeO___);
                exc.Write(p, 6, opShLoad[i].Fe2O3___);
                exc.Write(p, 7, opShLoad[i].SiO2___);
                exc.Write(p, 8, opShLoad[i].Al2O3___);
                exc.Write(p, 9, opShLoad[i].CaO___);
                exc.Write(p, 10, opShLoad[i].MgO___);
                exc.Write(p, 11, opShLoad[i].P___);
                exc.Write(p, 12, opShLoad[i].S___);
                exc.Write(p, 13, opShLoad[i].MnO___);
                exc.Write(p, 14, opShLoad[i].ZnO___);
                exc.Write(p, 15, opShLoad[i].PPP___);
                exc.Write(p, 16, opShLoad[i].H2O___);
                exc.Write(p, 17, opShLoad[i].TiO2___);
                exc.Write(p, 18, opShLoad[i].Cr___);
            }

            int lastFreeRow = opShLoad.Count+4;
            foreach (var p in himzolakok)
            {
                OptShihtDC_HimZolaKoks_Load1_Result himzola = new OptShihtDC_HimZolaKoks_Load1_Result();
                himzola.Descr = p.Descr;
                himzola.Al2O3 = p.Al2O3;
                himzola.CaO = p.CaO;
                himzola.Fe2O3 = p.Fe2O3;
                himzola.MgO = p.MgO;
                himzola.MnO = p.MnO;
                himzola.SiO2 = p.SiO2;
                himzola.TiO2 = p.TiO2;
                himzolaList.Add(himzola);
            }
            exc.Write(lastFreeRow+1,1,"Химический состав золы кокса");
            exc.Write(lastFreeRow+2,1,"Описание");
            exc.Write(lastFreeRow + 2, 2, "Al2O3, %");
            exc.Write(lastFreeRow + 2, 3, "CaO, %");
            exc.Write(lastFreeRow + 2, 4, "Fe2O3, %");
            exc.Write(lastFreeRow + 2, 5, "MgO, %");
            exc.Write(lastFreeRow + 2, 6, "MnO, %");
            exc.Write(lastFreeRow + 2, 7, "SiO2, %");
            exc.Write(lastFreeRow + 2, 8, "TiO2, %");
            for (int i = 0; i < himzolaList.Count; i++)
            {
                int p = lastFreeRow + 3+i;
                exc.Write(p,1,himzolaList[i].Descr);
                exc.Write(p, 2, himzolaList[i].Al2O3);
                exc.Write(p, 3, himzolaList[i].CaO);
                exc.Write(p, 4, himzolaList[i].Fe2O3);
                exc.Write(p, 5, himzolaList[i].MgO);
                exc.Write(p, 6, himzolaList[i].MnO);
                exc.Write(p, 7, himzolaList[i].SiO2);
                exc.Write(p, 8, himzolaList[i].TiO2);
            }
            test.Database.Connection.Close();
            exc.ChangeWorkSheet("Результат расчета");
            exc.Write(1,1,dataparameter);
            exc.Write(2,1,"Результаты расчетов");
            exc.Write(3,2,ResOkatSsgok);
            exc.Write(4, 2, ResOkatLebedinsk);
            exc.Write(5, 2, ResOkatKachanarsk);
            exc.Write(6, 2, ResOkatMihailovsk);
            exc.Write(7,1,"Параметры");
            exc.Write(8, 2, ResVyazkKonShlak);
            exc.Write(9, 2, ResOsnovKonShlak);
            exc.Write(10, 2, ResVyazkKonShlak14);
            exc.Write(11, 2, ResGradVyazkShlak);
            exc.Write(12, 2, ResAglomMmk);
            exc.SaveFile();
            exc.ExQuit();
            killDOZER("excel");
            Thread.Sleep(3000);
            string file = newpath;
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(file, contentType, Path.GetFileName(newpath));
        }

    }
}