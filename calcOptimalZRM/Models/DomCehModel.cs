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

    public class DomCehParameter
    {
        private double _proizvDomPechModel;

        public double proizvDomPechModel
        {
            get { return _proizvDomPechModel; }
            set { _proizvDomPechModel = value; }
        }

        private double _summUdRashKoksModel;
        public double summUdRashKoksModel
        {
            get { return _summUdRashKoksModel; }
            set { _summUdRashKoksModel = value; }
        }

        private double _sodSeraChugunModel;

        public double sodSeraChugunModel
        {
            get { return _sodSeraChugunModel; }
            set { _sodSeraChugunModel = value; }
        }

        private double _dolAglomeratMmkModel;
        public double dolAglomeratMmkModel
        {
            get { return _dolAglomeratMmkModel; }
            set { _dolAglomeratMmkModel = value; }
        }
        private double _Al2O3ShlakModel;
        public double Al2O3ShlakModel
        {
            get { return _Al2O3ShlakModel; }
            set { _Al2O3ShlakModel = value; }
        }
        private double _MgOShlakModel;
        public double MgOShlakModel
        {
            get { return _MgOShlakModel; }
            set { _MgOShlakModel = value; }
        }
    }

    public class VvodOgranicheniy
    {
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


    }

    public class ParamObj
    {
        private DateTime _pickDateTime;

        public DateTime pickDateTime
        {
            get { return _pickDateTime;}
            set { _pickDateTime = value; }
        }
        private byte _nPech;

        public byte nPech
        {
            get { return _nPech;}
            set { _nPech = value; }
        }
    }

    public class PechParams
    {
        public DateTime Pechdate { get; set; }
        public string Descr { get; set; }
        public float? Val { get; set; }
    }

    public class Okatishi
    {
        private double _okatSsgokModel;

        public double okatSsgokModel
        {
            get { return _okatSsgokModel; }
            set { _okatSsgokModel = value; }
        }

        private double _okatLebedinskModel;
        public double okatLebedinskModel
        {
            get { return _okatLebedinskModel; }
            set { _okatLebedinskModel = value; }
        }

        private double _okatKachanarskieModel;

        public double okatKachanarskieModel
        {
            get { return _okatKachanarskieModel; }
            set { _okatKachanarskieModel = value; }
        }

        private double _okatMihailovskModel;
        public double okatMihailovskModel
        {
            get { return _okatMihailovskModel; }
            set { _okatMihailovskModel = value; }
        }
    }

}