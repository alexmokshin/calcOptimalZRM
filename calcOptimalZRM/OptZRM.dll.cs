using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using calcOptimalZRM.Models;
using Microsoft.SolverFoundation.Common;
using System.IO;
using System.Threading;
using System.Web.Mvc;
using System.Web;
using System.Web.Configuration;
using Microsoft.Ajax.Utilities;

namespace calcOptimalZRM
{
    public class OptZRM
    {
        private DateTime _nachDateTime;
        public DateTime nachDataPech
        {
            get { return _nachDateTime; }
            set { _nachDateTime = value; }
        }

        private byte _nomerPech;
        public byte nomerPech
        {
            get { return _nomerPech; }
            set { _nomerPech = value; }
        }
        #region поля для окатышей
        private string _okatSsgok;

        public string okatSsgok
        {
            get { return _okatSsgok; }
            set { _okatSsgok = value; }
        }

        private string _okatLebedinsk;
        public string okatLebedinsk
        {
            get { return _okatLebedinsk; }
            set { _okatLebedinsk = value; }
        }

        private string _okatKachanarskie;

        public string okatKachanarskie
        {
            get { return _okatKachanarskie; }
            set { _okatKachanarskie = value; }
        }

        private string _okatMihailovsk;
        public string okatMihailovsk
        {
            get { return _okatMihailovsk; }
            set { _okatMihailovsk = value; }
        }
        #endregion
        #region поля для параметров плавки
        private string _proizvDomPechModel;

        public string proizvDomPechModel
        {
            get { return _proizvDomPechModel; }
            set { _proizvDomPechModel = value; }
        }

        private string _summUdRashKoksModel;
        public string summUdRashKoksModel
        {
            get { return _summUdRashKoksModel; }
            set { _summUdRashKoksModel = value; }
        }

        private string _sodSeraChugunModel;

        public string sodSeraChugunModel
        {
            get { return _sodSeraChugunModel; }
            set { _sodSeraChugunModel = value; }
        }

        private string _dolAglomeratMmkModel;
        public string dolAglomeratMmkModel
        {
            get { return _dolAglomeratMmkModel; }
            set { _dolAglomeratMmkModel = value; }
        }
        private string _Al2O3ShlakModel;
        public string Al2O3ShlakModel
        {
            get { return _Al2O3ShlakModel; }
            set { _Al2O3ShlakModel = value; }
        }
        private string _MgOShlakModel;
        public string MgOShlakModel
        {
            get { return _MgOShlakModel; }
            set { _MgOShlakModel = value; }
        }
        #endregion
        #region Ограничения
        public double? vzskkonshlak { get => _vzskkonshlak; set => _vzskkonshlak = value; }
        public double? vzskkonshlakmin { get => _vzskkonshlakmin; set => _vzskkonshlakmin = value; }
        public double? vzskkonshlakmax { get => _vzskkonshlakmax; set => _vzskkonshlakmax = value; }
        
        public double? osnkonshlack { get => _osnkonshlack; set => _osnkonshlack = value; }
        public double? onskonshlakmin { get => _onskonshlakmin; set => _onskonshlakmin = value; }
        public double? osnkonshlakmax { get => _osnkonshlakmax; set => _osnkonshlakmax = value; }
        public double? onskonshlakkzad { get => _onskonshlakkzad; set => _onskonshlakkzad = value; }
        public double? vzskkonshlak1450 { get => _vzskkonshlak1450; set => _vzskkonshlak1450 = value; }
        public double? vzskkonshlakmin1450 { get => _vzskkonshlakmin1450; set => _vzskkonshlakmin1450 = value; }
        public double? vzskkonshlakmax1450 { get => _vzskkonshlakmax1450; set => _vzskkonshlakmax1450 = value; }
        
        public double? grdvzkkkonshlak { get => _grdvzkkkonshlak; set => _grdvzkkkonshlak = value; }
        public double? grdvzkkkonshlakmin { get => _grdvzkkkonshlakmin; set => _grdvzkkkonshlakmin = value; }
        public double? grdvzkkkonshlakmax { get => _grdvzkkkonshlakmax; set => _grdvzkkkonshlakmax = value; }
       
        public double? aglmmknow { get => _aglmmknow; set => _aglmmknow = value; }
        public double? aglmmkmin { get => _aglmmkmin; set => _aglmmkmin = value; }
        public double? aglmmkmax { get => _aglmmkmax; set => _aglmmkmax = value; }
       
        public double? summdoleyzrmnow { get => _summdoleyzrmnow; set => _summdoleyzrmnow = value; }

        

        
        private double? _vzskkonshlak;
        private double? _vzskkonshlakmin;
        private double? _vzskkonshlakmax;
        private double? _osnkonshlack;
        
        private double? _onskonshlakmin;
        private double? _osnkonshlakmax;
        private double? _onskonshlakkzad;
        private double? _vzskkonshlak1450;
        private double? _vzskkonshlakmin1450;
        private double? _vzskkonshlakmax1450;
        
        private double? _grdvzkkkonshlak;
        private double? _grdvzkkkonshlakmin;
        private double? _grdvzkkkonshlakmax;
        
        private double? _aglmmknow;
        private double? _aglmmkmin;
        private double? _aglmmkmax;
       
        private double? _summdoleyzrmnow;
        #endregion
        public static string path;

        public void SetDataOnPechInExcel(string excelPath)
        {
            DateTime dt = nachDataPech;
            int npech = nomerPech;
            path = excelPath;
            ExcelReport Excel = new ExcelReport(path, false);
            List<OptShihtDC_Shihta_Load_Result> opShLoad = new List<OptShihtDC_Shihta_Load_Result>();
            List<OptShihtDC_HimZolaKoks_Load1_Result> himzolaList = new List<OptShihtDC_HimZolaKoks_Load1_Result>();
            List<PechParams> ppList = new List<PechParams>();
            testReportEntities test = new testReportEntities();
            test.Database.Connection.Open();
            ParamObj pobj = new ParamObj();
            DateTime pechDateTime = pobj.pickDateTime;
            byte pech = pobj.nPech;
            //получить химсостав золы и кокса
            var shihtaLoad = test.OptShihtDC_Shihta_Load(dt, Convert.ToByte(npech), 0, 0);
            var paramplavk = test.OptShihtDC_Param_Load(dt, 0);
            var himzolakok = test.OptShihtDC_HimZolaKoks_Load1(dt, 0);
            #region наполнение листа параметрами золы кокса из базы

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
            #endregion
            #region Наполнение листа параметрами шихты из базы
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
            #endregion
            #region Наполнение листа параметрами плавки из базы
            foreach (var p in paramplavk)
            {
                PechParams pp = new PechParams();
                if (npech == 1)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP1;
                }
                if (npech == 2)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP2;
                }
                if (npech == 3)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP3;
                }
                if (npech == 4)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP4;
                }
                if (npech == 5)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP5;
                }
                if (npech == 6)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP6;
                }
                if (npech == 7)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP7;
                }
                if (npech == 8)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP8;
                }
                if (npech == 9)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP9;
                }
                if (npech == 10)
                {
                    pp.Pechdate = p.dtFirstDay;
                    pp.Descr = p.Descr;
                    pp.Val = p.sDP10;
                }
                ppList.Add(pp);
            }
            #endregion
            #region Загрузка в Excel значений по параметрам печи

            double value;
            Excel.ChangeWorkSheet("Исходные данные и результаты");
            foreach (var p in ppList)
            {
                value = Convert.ToDouble(p.Val);
                if (p.Descr == "Суточная производительность доменной печи, т/сут.")
                {
                    Excel.Write(3, 2, Math.Round(value, 2));
                }

                if (p.Descr == "Удельный выход шлака, кг/т чугуна")
                {
                    Excel.Write(4, 2, Math.Round(value, 2));
                }

                if (p.Descr == "Выход колошниковой пыли, кг/т чуг.")
                {
                    Excel.Write(6, 2, Math.Round(value, 2));
                    Excel.Write(6, 3, Math.Round(value, 2));
                }

                if (p.Descr == "Давление дутья, ати")
                {
                    Excel.Write(8, 2, Math.Round(value, 2));
                    Excel.Write(8, 3, Math.Round(value, 2));
                }

                if (p.Descr == "Температура дутья, °С")
                {
                    Excel.Write(9, 2, Math.Round(value, 0));
                    Excel.Write(9, 3, Math.Round(value, 0));
                }

                if (p.Descr == "Влажность дутья, г/м3")
                {
                    Excel.Write(10, 2, Math.Round(value, 3));
                    Excel.Write(10, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание кислорода в дутье, %")
                {
                    Excel.Write(11, 2, Math.Round(value, 3));
                    Excel.Write(11, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Расход природного газа, м3/т чугуна")
                {
                    Excel.Write(12, 2, Math.Round(value, 3));
                    Excel.Write(12, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание кремния в чугуне, [Si], % (масс.)")
                {
                    Excel.Write(14, 2, Math.Round(value, 3));
                    Excel.Write(14, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание серы в чугуне, [S], % (масс.)")
                {
                    Excel.Write(15, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание марганца в чугуне, [Mn], % (масс.)")
                {
                    Excel.Write(16, 2, Math.Round(value, 3));
                    Excel.Write(16, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание углерода в чугуне, [C], % (масс.)")
                {
                    Excel.Write(17, 2, Math.Round(value, 3));
                    Excel.Write(17, 3, Math.Round(value, 3));
                }

                if (p.Descr == "Содержание фосфора в чугуне, [P], % (масс.)")
                {
                    Excel.Write(18, 2, Math.Round(value, 3));
                    Excel.Write(18, 3, Math.Round(value, 3));
                }
                if (p.Descr == "Содержание титана в чугуне, [Ti], % (масс.)")
                {
                    Excel.Write(19, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: CaO, %")
                {
                    Excel.Write(21, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: SiO2, %")
                {
                    Excel.Write(22, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: Al2O3, %")
                {
                    Excel.Write(23, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: MgO, %")
                {
                    Excel.Write(24, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: S, %")
                {
                    Excel.Write(25, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Химический состав шлака: TiO2, %")
                {
                    Excel.Write(26, 2, Math.Round(value, 3));
                }

                if (p.Descr == "Удельный расход кокса, кг/т чугуна")
                {
                    Excel.Write(74, 2, Math.Round(value, 3));
                }
            }
            Excel.Write(75, 2, "1470");
            Excel.Write(75, 3, "1470");
            //Excel.SaveFile();
            Excel.ChangeWorkSheet("Ввод составов (база)");
            foreach (var pp in ppList)
            {
                if (pp.Descr == "Зола кокса(А), %")
                {
                    Excel.Write(33, 1, Math.Round(Convert.ToDouble(pp.Val), 3));
                }

                if (pp.Descr == "Сера кокса (S), %")
                {
                    Excel.Write(33, 2, Math.Round(Convert.ToDouble(pp.Val), 3));
                }

                if (pp.Descr == "Летучие кокса (Л), %")
                {
                    Excel.Write(33, 3, Math.Round(Convert.ToDouble(pp.Val), 3));
                }
            }
            //Excel.SaveFile();
            #endregion
            #region Загрузка в Excel значений по ЖРМ
            Excel.ChangeWorkSheet("Исходные данные и результаты");
            Excel.Write(29, 2, 0);
            Excel.Write(30, 2, 0);
            Excel.Write(31, 2, 0);
            Excel.Write(32, 2, 0);
            Excel.Write(33, 2, 0);
            Excel.Write(35, 2, 0);
            Excel.Write(36, 2, 0);
            Excel.Write(37, 2, 0);
            Excel.Write(38, 2, 0);
            Excel.Write(39, 2, 0);
            Excel.Write(40, 2, 0);
            Excel.Write(41, 2, 0);
            Excel.Write(42, 2, 0);
            Excel.Write(43, 2, 0);
            Excel.Write(44, 2, 0);
            Excel.Write(45, 2, 0);
            Excel.Write(46, 2, 0);
            foreach (OptShihtDC_Shihta_Load_Result currMaterialGRM in opShLoad)
            {
                if (currMaterialGRM.Материал == "Агл фаб 2")
                {
                    Excel.Write(29, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Агл фаб 3")
                {
                    Excel.Write(30, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Агл фаб 4")
                {
                    Excel.Write(31, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Агломерат неочищенный")
                {
                    Excel.Write(32, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Агл яма")
                {
                    Excel.Write(33, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Окат.Сокол неофлюсов.")
                {
                    Excel.Write(35, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Окат.Лебед")
                {
                    Excel.Write(36, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Окат.Качкан")
                {
                    Excel.Write(37, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Окат.Михайл")
                {
                    Excel.Write(38, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Шлак Свароч")
                {
                    Excel.Write(39, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Королек")
                {
                    Excel.Write(40, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Домен.Присад")
                {
                    Excel.Write(41, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Руда мих. доменная")
                {
                    Excel.Write(42, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Руда Марг Жайремская")
                {
                    Excel.Write(43, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }
                if (currMaterialGRM.Материал == "Агломерат мелкий")
                {
                    Excel.Write(44, 2, Math.Round(Convert.ToDouble(currMaterialGRM.Расход__кг_т_чугуна), 2));
                }

            }
            Excel.Write(59, 3, Math.Round(double.Parse(Excel.GetValue("B59").ToString()), 4));
            Excel.Write(60, 3, Math.Round(double.Parse(Excel.GetValue("B60").ToString()), 4));
            Excel.Write(61, 3, Math.Round(double.Parse(Excel.GetValue("B61").ToString()), 4));
            Excel.Write(62, 3, Math.Round(double.Parse(Excel.GetValue("B62").ToString()), 4));
            Excel.Write(63, 3, Math.Round(double.Parse(Excel.GetValue("B63").ToString()), 4));
            Excel.Write(64, 3, Math.Round(double.Parse(Excel.GetValue("B64").ToString()), 4));
            Excel.Write(65, 3, Math.Round(double.Parse(Excel.GetValue("B65").ToString()), 4));
            Excel.Write(66, 3, Math.Round(double.Parse(Excel.GetValue("B66").ToString()), 4));

            double _dol_agloMMK = Math.Round(double.Parse(Excel.GetValue("B54").ToString()), 4);
            double _dol_okatSokol = Math.Round(double.Parse(Excel.GetValue("B55").ToString()), 4);
            double _dol_okatLebed = Math.Round(double.Parse(Excel.GetValue("B56").ToString()), 4);
            double _dol_okatKatch = Math.Round(double.Parse(Excel.GetValue("B57").ToString()), 4);
            double _dol_okatMih = Math.Round(double.Parse(Excel.GetValue("B58").ToString()), 4);
            Excel.ChangeWorkSheet("Соотношение расходов ЖРМ"); // лист "Соотношение расходов ЖРМ"
            Excel.Write(8, 3, _dol_agloMMK);
            Excel.Write(29, 2, _dol_okatSokol);
            Excel.Write(30, 2, _dol_okatLebed);
            Excel.Write(31, 2, _dol_okatKatch);
            Excel.Write(32, 2, _dol_okatMih);
            //Excel.SaveFile();
            #endregion
            #region Ввод составов (база)
            Excel.ChangeWorkSheet("Ввод составов (база)");
            foreach (var currMaterialGRM in opShLoad)
            {
                if (currMaterialGRM.Материал == "Агл фаб 2")
                {
                    Excel.Write(4, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(4, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(4, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(4, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(4, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(4, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(4, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(4, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(4, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(4, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(4, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агл фаб 3")
                {
                    Excel.Write(5, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(5, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(5, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(5, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(5, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(5, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(5, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(5, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(5, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(5, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(5, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агл фаб 4")
                {
                    Excel.Write(6, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(6, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(6, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(6, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(6, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(6, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(6, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(6, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(6, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(6, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(6, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агломерат неочищенный")
                {
                    Excel.Write(7, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(7, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(7, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(7, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(7, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(7, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(7, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(7, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(7, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(7, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(7, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агл яма")
                {
                    Excel.Write(8, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(8, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(8, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(8, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(8, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(8, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(8, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(8, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(8, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(8, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(8, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Сокол неофлюсов.")
                {
                    Excel.Write(9, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(9, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(9, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(9, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(9, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(9, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(9, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(9, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(9, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(9, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(9, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Лебед")
                {
                    Excel.Write(10, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(10, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(10, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(10, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(10, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(10, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(10, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(10, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(10, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(10, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(10, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Качкан")
                {
                    Excel.Write(11, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(11, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(11, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(11, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(11, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(11, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(11, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(11, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(11, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(11, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(11, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Михайл")
                {
                    Excel.Write(12, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(12, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(12, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(12, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(12, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(12, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(12, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(12, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(12, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(12, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(12, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Шлак Свароч")
                {
                    Excel.Write(13, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(13, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(13, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(13, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(13, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(13, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(13, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(13, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(13, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(13, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(13, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Королек")
                {
                    Excel.Write(14, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(14, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(14, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(14, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(14, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(14, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(14, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(14, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(14, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(14, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(14, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Домен.Присад")
                {
                    Excel.Write(15, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(15, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(15, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(15, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(15, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(15, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(15, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(15, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(15, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(15, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(15, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Руда мих. доменная")
                {
                    Excel.Write(16, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(16, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(16, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(16, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(16, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(16, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(16, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(16, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(16, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(16, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(16, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Руда Марг Жайремская")
                {
                    Excel.Write(17, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(17, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(17, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(17, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(17, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(17, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(17, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(17, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(17, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(17, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(17, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агломерат мелкий")
                {
                    Excel.Write(18, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(18, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(18, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(18, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(18, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(18, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(18, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(18, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(18, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(18, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(18, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
            }
            #endregion
            #region Ввод составов (база) состав кокса

            foreach (var p in himzolaList)
            {
                Excel.Write(37, 1, Math.Round(p.Fe2O3, 3));
                Excel.Write(37, 2, Math.Round(p.CaO, 3));
                Excel.Write(37, 3, Math.Round(p.SiO2, 3));
                Excel.Write(37, 4, Math.Round(p.Al2O3, 3));
                Excel.Write(37, 5, Math.Round(p.MgO, 3));
                Excel.Write(37, 6, 0);
            }
            //Excel.SaveFile();
            #endregion
            #region Ввод составов (проект) ЖРМ
            Excel.ChangeWorkSheet("Ввод составов (проект)");
            foreach (var currMaterialGRM in opShLoad)
            {
                if (currMaterialGRM.Материал == "Окат.Сокол неофлюсов.")
                {
                    Excel.Write(5, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(5, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(5, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(5, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(5, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(5, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(5, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(5, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(5, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(5, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(5, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Лебед")
                {
                    Excel.Write(6, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(6, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(6, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(6, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(6, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(6, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(6, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(6, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(6, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(6, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(6, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Качкан")
                {
                    Excel.Write(7, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(7, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(7, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(7, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(7, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(7, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(7, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(7, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(7, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(7, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(7, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Окат.Михайл")
                {
                    Excel.Write(8, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(8, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(8, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(8, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(8, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(8, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(8, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(8, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(8, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(8, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(8, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Шлак Свароч")
                {
                    Excel.Write(9, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(9, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(9, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(9, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(9, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(9, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(9, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(9, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(9, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(9, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(9, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Королек")
                {
                    Excel.Write(10, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(10, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(10, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(10, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(10, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(10, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(10, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(10, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(10, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(10, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(10, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Домен.Присад")
                {
                    Excel.Write(11, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(11, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(11, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(11, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(11, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(11, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(11, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(11, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(11, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(11, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(11, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Руда мих. доменная")
                {
                    Excel.Write(12, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(12, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(12, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(12, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(12, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(12, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(12, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(12, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(12, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(12, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(12, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Руда Марг Жайремская")
                {
                    Excel.Write(13, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(13, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(13, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(13, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(13, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(13, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(13, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(13, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(13, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(13, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(13, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
                if (currMaterialGRM.Материал == "Агломерат мелкий")
                {
                    Excel.Write(14, 3, Math.Round(Convert.ToDouble(currMaterialGRM.Fe___), 3));
                    Excel.Write(14, 4, Math.Round(Convert.ToDouble(currMaterialGRM.FeO___), 3));
                    Excel.Write(14, 5, Math.Round(Convert.ToDouble(currMaterialGRM.Fe2O3___), 3));
                    Excel.Write(14, 6, Math.Round(Convert.ToDouble(currMaterialGRM.SiO2___), 3));
                    Excel.Write(14, 7, Math.Round(Convert.ToDouble(currMaterialGRM.Al2O3___), 3));
                    Excel.Write(14, 8, Math.Round(Convert.ToDouble(currMaterialGRM.CaO___), 3));
                    Excel.Write(14, 9, Math.Round(Convert.ToDouble(currMaterialGRM.MgO___), 3));
                    Excel.Write(14, 10, Math.Round(Convert.ToDouble(currMaterialGRM.P___), 3));
                    Excel.Write(14, 11, Math.Round(Convert.ToDouble(currMaterialGRM.S___), 3));
                    Excel.Write(14, 12, Math.Round(Convert.ToDouble(currMaterialGRM.MnO___), 3));
                    Excel.Write(14, 13, Math.Round(Convert.ToDouble(currMaterialGRM.Cr___), 3));
                }
            }
            #endregion
            #region ввод составов (проект) Технический состав золы кокса
            foreach (var pp in ppList)
            {
                if (pp.Descr == "Зола кокса(А), %")
                {
                    Excel.Write(28, 1, Math.Round(Convert.ToDouble(pp.Val), 3));
                }

                if (pp.Descr == "Сера кокса (S), %")
                {
                    Excel.Write(28, 2, Math.Round(Convert.ToDouble(pp.Val), 3));
                }

                if (pp.Descr == "Летучие кокса (Л), %")
                {
                    Excel.Write(28, 3, Math.Round(Convert.ToDouble(pp.Val), 3));
                }
            }
            #endregion
            #region Ввод составов (проект) состав золы кокса
            foreach (var p in himzolaList)
            {
                Excel.Write(32, 1, Math.Round(p.Fe2O3, 3));
                Excel.Write(32, 2, Math.Round(p.CaO, 3));
                Excel.Write(32, 3, Math.Round(p.SiO2, 3));
                Excel.Write(32, 4, Math.Round(p.Al2O3, 3));
                Excel.Write(32, 5, Math.Round(p.MgO, 3));
                Excel.Write(32, 6, 0);
            }
            #endregion
            Thread.Sleep(5000);
            Excel.SaveFile();
            Excel.ChangeWorkSheet("Соотношение расходов ЖРМ");
            proizvDomPechModel = Excel.GetValue("B5");
            summUdRashKoksModel = Excel.GetValue("B6");
            sodSeraChugunModel = Excel.GetValue("B7");
            dolAglomeratMmkModel = Excel.GetValue("B8");
            Al2O3ShlakModel = Excel.GetValue("B9");
            MgOShlakModel = Excel.GetValue("B10");
            vzskkonshlak = GetTroolyExcel("B17", Excel);
            vzskkonshlakmin = GetTroolyExcel("C17", Excel); 
            vzskkonshlakmax = GetTroolyExcel("D17", Excel); 
            
            osnkonshlack = GetTroolyExcel("B18", Excel);
            onskonshlakmin = GetTroolyExcel("C18", Excel); 
            osnkonshlakmax = GetTroolyExcel("D18", Excel); 
            onskonshlakkzad = GetTroolyExcel("E18", Excel); 
            vzskkonshlak1450 = GetTroolyExcel("B19", Excel); 
            vzskkonshlakmin1450 = GetTroolyExcel("C19", Excel); 
            vzskkonshlakmax1450 = GetTroolyExcel("D19", Excel); 
            
            grdvzkkkonshlak = GetTroolyExcel("B20", Excel); 
            grdvzkkkonshlakmin = GetTroolyExcel("C20", Excel); 
            grdvzkkkonshlakmax = GetTroolyExcel("D20", Excel);
            
            aglmmknow = GetTroolyExcel("B21", Excel);
            aglmmkmin = GetTroolyExcel("C21", Excel);
            aglmmkmax = GetTroolyExcel("D21", Excel);
           
            summdoleyzrmnow = GetTroolyExcel("B22", Excel);
            Excel.ExQuit();
            test.Database.Connection.Close();

        }

        public void CalcOptZRM(bool[] constraint, int target, int influence, string excPath)
        {
            ExcelReport exc = new ExcelReport(excPath, true);
            exc.ChangeWorkSheet("Соотношение расходов ЖРМ");
            for (int i = 0; i < constraint.Length; i++)
            {
                if (constraint[i] == false)
                {
                    exc.Write(17 + i, 7, false);
                }
                else
                {
                    exc.Write(17 + i, 7, true);
                }
            }
            exc.Write(22, 7, target);
            exc.Write(33, 7, influence);
            //exc.Write(29, 2, 0);
            //exc.Write(30, 2, 0);
            //exc.Write(31, 2, 0);
            //exc.Write(32, 2, 0);
            exc.RunMacs("ОптимизацияСоотношениеЖРМ");
            Thread.Sleep(5000);
            okatSsgok = exc.GetValue("B29");
            okatLebedinsk = exc.GetValue("B30");
            okatKachanarskie = exc.GetValue("B31");
            okatMihailovsk = exc.GetValue("B32");
            vzskkonshlak = GetTroolyExcel("B17",exc);
            osnkonshlack = GetTroolyExcel("B18", exc);
            vzskkonshlak1450 = GetTroolyExcel("B19", exc);
            grdvzkkkonshlak = GetTroolyExcel("B20", exc);
            aglmmknow = GetTroolyExcel("B21", exc);
            summdoleyzrmnow = GetTroolyExcel("B22", exc);
            //exc.SaveFile();
            exc.ExQuit(false);


        }

        public double? GetTroolyExcel(string range, ExcelReport exc)
        {
            double? d;
            string s = exc.GetValue(range);
            try
            {
                d = Convert.ToDouble(s);
            }
            catch (Exception e)
            {
                d = null;
                
            }

            return d;


        }
    }



}