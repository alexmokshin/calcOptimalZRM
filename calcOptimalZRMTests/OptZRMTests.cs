using Microsoft.VisualStudio.TestTools.UnitTesting;
using calcOptimalZRM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace calcOptimalZRM.Tests
{
    [TestClass()]
    public class OptZRMTests
    {
        [TestMethod()]
        public void SetDataOnPechInExcelTest()
        {
            OptZRM opz = new OptZRM();
            opz.nachDataPech = Convert.ToDateTime("01.12.2014");
            opz.nomerPech = 1;
            string path =
                @"C:\Users\Александр\Documents\учебка\диплом\web\calcOptimalZRM\calcOptimalZRM\Content\Оптимальная доменная шихта 2010_.xlt";
            opz.SetDataOnPechInExcel(path);
            Thread.Sleep(5000);
            ExcelReport exp = new ExcelReport(path,true);
            exp.ChangeWorkSheet("Ввод составов (база)");
            string g = exp.GetValue("B7");
            double c = Convert.ToDouble(g);
            Assert.AreEqual(12.80,c,1);
        }
    }
}