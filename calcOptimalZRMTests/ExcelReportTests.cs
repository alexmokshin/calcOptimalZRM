using Microsoft.VisualStudio.TestTools.UnitTesting;
using calcOptimalZRM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace calcOptimalZRM.Tests
{
    [TestClass()]
    public class ExcelReportTests
    {
        [TestMethod()]
        public void RunMacTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ChangeWorkSheetTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ChangeWorkSheetTest1()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ExcelReportTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void MergeTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void WriteTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void SaveTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void GetValueTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void SetValueTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void RunMacsTest()
        {
            string path = @"D:/testexcel1.xlsm";
            ExcelReport exp = new ExcelReport(path,true);
            exp.ChangeWorkSheet("Соотношение расходов ЖРМ");
            exp.Write(29, 2, 0);
            exp.Write(30, 2, 0);
            exp.Write(31, 2, 0);
            exp.Write(32, 2, 0);
            exp.RunMacs("ОптимизацияСоотношениеЖРМ");
            Thread.Sleep(5000);
            string d = exp.GetValue("B29");
            double m = Double.Parse(d);
            Assert.IsTrue(m > 0);
            exp.Write(34,2,1000);
            exp.Write(33, 2, 100);
            Console.WriteLine("Результат расчета: {0}", m);
            //exp.SaveFile();
            exp.Write(35,2,1000);
            exp.Write(36,2,1000);
            //exp.SaveFile();
            //exp.ExQuit();
        }
    }
}