using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;
using calcOptimalZRM.Models;
using System.Linq;
using System.Data.Entity;
using System.Net.Mime;


namespace calcOptimalZRM.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string time = "2014-12-01";
            DateTime truetime = Convert.ToDateTime(time);
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
            testReportEntities test = new testReportEntities();
            var shihtaLoad = test.OptShihtDC_Shihta_Load(Convert.ToDateTime("2014-12-01"), 1, 0, 0);
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
            testReportEntities test = new testReportEntities();
            var query = from dnp in test.DC_NSI_Pech where dnp.actual == true select dnp;
            foreach (var result in query)
            {
                actPech.Add(new SelectListItem {Text = result.Name, Value = Convert.ToString(result.PechId)});
            }

            ViewBag.PechType = actPech;
            List<SelectListItem> months = new List<SelectListItem>
            {
                new SelectListItem {Text = "Январь", Value = "1"},
                new SelectListItem {Text = "Февраль", Value = "2"},
                new SelectListItem {Text = "Март", Value = "3"},
                new SelectListItem {Text = "Апрель", Value = "4"},
                new SelectListItem {Text = "Май", Value = "5"},
                new SelectListItem {Text = "Июнь", Value = "6"},
                new SelectListItem {Text = "Июль", Value = "7"},
                new SelectListItem {Text = "Август", Value = "8"},
                new SelectListItem {Text = "Сентябрь", Value = "9"},
                new SelectListItem {Text = "Октябрь", Value = "10"},
                new SelectListItem {Text = "Ноябрь", Value = "11"},
                new SelectListItem {Text = "Декабрь", Value = "12"}
            };
            ViewBag.Monthes = months;
            return View();
        }
        [HttpPost]
        public ActionResult CalcZHRM(int? Monthes, int? PechType, string YearPick)
        {
            
            return View();

        }
        [HttpGet]
        public ActionResult GetDateAndPech(List<string> names)
        {
            // for (int i = 0; i < names.L)
            //string res = names
            // return View(CalcZHRM());
            return RedirectToAction("CalcZHRM");
        }

    }
}