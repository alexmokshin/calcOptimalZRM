using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace calcOptimalZRM.Models
{
    public class DomCehModel
    {
        //System.Data.SqlClient.SqlParameter param = new System.Data.SqlClient.SqlParameter("@dtFirstDay", "2014-01-12");
        //testReportEntities trewq = new testReportEntities();
        //testReportEntities GetEntities = new testReportEntities();
        [DataType(DataType.Date)]
        public DateTime BirthDate { get; set; }

    }
    public class ShihtaLoad
    {
        public DateTime dtFirstDay { get; set; }
        public byte Номер_печи { get; set; }
        public short Код_материала { get; set; }
        public byte Тип_материала { get; set; }
        public string Материал { get; set; }
        public decimal Расход__кг_т_чугуна { get; set; }
        public decimal Доля { get; set; }
        public decimal Fe___ { get; set; }
        public decimal FeO___ { get; set; }
        public decimal Fe2O3___ { get; set; }
        public decimal SiO2___ { get; set; }
        public decimal Al2O3___ { get; set; }
        public decimal CaO___ { get; set; }
        public decimal MgO___ { get; set; }
        public decimal P___ { get; set; }
        public decimal S___ { get; set; }
        public decimal MnO___ { get; set; }
        public decimal ZnO___ { get; set; }
        public decimal PPP___ { get; set; }
        public decimal H2O___ { get; set; }
        public decimal TiO2___ { get; set; }
        public decimal Cr___ { get; set; }
    }

}