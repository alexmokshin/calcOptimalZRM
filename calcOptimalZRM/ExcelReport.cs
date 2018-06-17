using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Reflection;

namespace calcOptimalZRM
{
    public class ExcelReport
    {

        //private Timer timer;
        //private ToolStripProgressBar progress;

        private enum COLROW { ROW = 0, COL };

        private Application _excelApp;
        private Worksheet _workSheet;
        private Workbook _workBook;

        private Int32 _intLastColWrite;
        private Int32 _intLastRowWrite;

        private Boolean _bBorderAround;
        private Boolean _bMergeOnlyRows = true;
        private Int32 _intPrecision;

        private List<Int32> _intLstColNoShow = new List<Int32>();
        private List<Int32> _intLstRowNoShow = new List<Int32>();

        object Range;

        // Логическая переменная для отображения границ ячеек при выводе в Excel
        public Boolean BorderAround
        {
            get
            {
                return _bBorderAround;
            }
            set
            {
                _bBorderAround = value;
            }
        }
        // Переменная для настройки количества десятичных знаков при выводе числел в ячейку Excel
        public Int32 Precision
        {
            get
            {
                return _intPrecision;
            }
            set
            {
                _intPrecision = value;
            }
        }

        public Application ExcelApp
        {
            get
            {
                return _excelApp;
            }
        }

        public Workbook WorkBook
        {
            get
            {
                return _workBook;
            }
        }

        public Worksheet WorkSheet
        {
            get
            {
                return _workSheet;
            }
        }

        public Int32 LastColWrite
        {
            get
            {
                return _intLastColWrite;
            }
        }

        public Int32 LastRowWrite
        {
            get
            {
                return _intLastRowWrite;
            }
        }

        public List<Int32> ColumnsNoShow
        {
            get
            {
                return _intLstColNoShow;
            }
            set
            {
                _intLstColNoShow = value;
            }
        }

        public List<Int32> RowsNoShow
        {
            get
            {
                return _intLstRowNoShow;
            }
            set
            {
                _intLstRowNoShow = value;
            }
        }

        public Boolean MergeRowsOnly   //если true, то при объединении диапазона объединяются только строки
        {
            get
            {
                return _bMergeOnlyRows;
            }
            set
            {
                _bMergeOnlyRows = value;
            }
        }


      
        public void ChangeWorkSheet(Int32 inPageNum)
        {
            this._workSheet = (Worksheet)_workBook.Worksheets[inPageNum];
        }

        public void ChangeWorkSheet(String strPageName)
        {
            this._excelApp.Visible = false;
            this._workSheet = (Worksheet)_workBook.Worksheets[strPageName];
        }

        // Конструктор
        public ExcelReport(String inPath, bool readOnly)
        {
            FileInfo fi = new FileInfo(inPath);
            if (!fi.Exists)
            {
                throw new ArgumentException(String.Format("Файл не найден: '{0}'", inPath));
            }
            inPath = fi.FullName;

            this._excelApp = new Application();
            this._excelApp.Visible = false;
            //Filename UpdateLinks, ReadOnly, Format,         Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad
            if (readOnly == false)
            {
                this._workBook =
                    _excelApp.Workbooks
                        .Open(inPath); //, (Object)0, (Object)false,  (Object)String.Format, (Object)String.Empty, (Object)String.Empty, (Object)true, (Object)null,   (Object)String.Empty, (Object)false, (Object)true, (Object)null, (Object)false, (Object)false, (Object)false);
            }
            else
            {
                this._workBook =
                    _excelApp.Workbooks
                        .Open(inPath,null,true); //, (Object)0, (Object)false,  (Object)String.Format, (Object)String.Empty, (Object)String.Empty, (Object)true, (Object)null,   (Object)String.Empty, (Object)false, (Object)true, (Object)null, (Object)false, (Object)false, (Object)false);

            }

            //.Workbooks.Add(Type.Missing);
            ChangeWorkSheet(1);
            this._excelApp.Visible = false;
        }

        // Деструктор
        ~ExcelReport()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this._workSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this._workBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this._excelApp);
            _workSheet = null;
            _workBook = null;
            _excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        //объединение ячеек
        public Boolean Merge(Int32 inRowLeft, Int32 inColLeft, Int32 inRowRight, Int32 inColRight)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)_workSheet.get_Range((Microsoft.Office.Interop.Excel.Range)_workSheet.Cells[inRowLeft, inColLeft], (Microsoft.Office.Interop.Excel.Range)_workSheet.Cells[inRowRight, inColRight]);
                rng.Merge(this.MergeRowsOnly);
                // Установить границу объединенных ячеек
                rng.BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void Write(Int32 inRowIndex, Int32 inColumnIndex, Object inValue)
        {
            this._excelApp.Visible = false;
            ((Range)_workSheet.Cells[inRowIndex, inColumnIndex]).Value2 = inValue;
        }
       

        private Boolean IsFloat(String inFloat, out String outFloat) //данная функция определяет являетсяли строковая переменная числом
        {
            Boolean result = true;
            Int32 intPoint = 0;
            String strTemp = String.Empty;
            Int32 intTemp = 0;
            outFloat = inFloat.Trim();

            // Проверить системные настройки разделителя десятичных знаков на компьютере пользователя
            // MessageBox.Show(System.Globalization.NumberFormatInfo.CurrentInfo.PercentDecimalSeparator.ToString(), "Сообщение", buttons, MessageBoxIcon.Information);
            // System.Globalization.NumberStyles.AllowDecimalPoint

            foreach (Char chr in outFloat)
            {
                if (chr == '-') continue;
                //if (chr == char.Parse(System.Globalization.NumberFormatInfo.CurrentInfo.PercentDecimalSeparator.ToString()))
                if ((chr == '.') || (chr == ','))
                {
                    if (intPoint > 1)
                    {
                        result = false;
                        return result;
                    }
                    intPoint = 1;
                    continue;
                }
                strTemp = chr.ToString();
                if (!int.TryParse(strTemp, out intTemp))
                {
                    result = false;
                    return result;
                }
            }
            if (result)
            {
                Char chrCurrentPercentDecimalSeparator;
                chrCurrentPercentDecimalSeparator = Char.Parse(System.Globalization.NumberFormatInfo.CurrentInfo.PercentDecimalSeparator.ToString());
                outFloat = inFloat.Trim().Replace('.', chrCurrentPercentDecimalSeparator);
                //strFloatOut = strFloatIn.Trim().Replace('.', ',');
            }
            return result;
        }


        // Метод сохранения таблицы [dtCurrent] в листе под номером intWorkSheetNumber, с указанным отступом
        public Boolean Save(System.Data.DataTable inTable, Int32 inWorkSheetNumber, Int32 inShiftHoriz, Int32 inShiftVert,
                            Int32 inIntervalHoriz, Int32 inIntervalVert)
        //                  горизонтальный интервал        вертикальный интервал
        {
            #region проверки
            if (inWorkSheetNumber < 1) return false;
            if (inShiftHoriz < 0) inShiftHoriz = 0;
            if (inShiftVert < 0) inShiftVert = 0;
            if (inIntervalHoriz < 0) inIntervalHoriz = 0;
            if (inIntervalVert < 0) inIntervalVert = 0;
            #endregion

            try
            {
                if (inWorkSheetNumber > this._excelApp.Worksheets.Count)
                {
                    _workSheet = (Worksheet)this._excelApp.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                else
                {
                    _workSheet = (Worksheet)this._excelApp.Worksheets[(int)inWorkSheetNumber];
                }

                Int32 intCurrCol = 1;
                Int32 intCurrRow = 1 + inShiftVert;

                foreach (DataRow dRowCurrent in inTable.Rows)
                {
                    if (!IsShow(COLROW.ROW, dRowCurrent.Table.Rows.IndexOf(dRowCurrent))) continue;
                    intCurrCol = 1 + inShiftHoriz;
                    foreach (DataColumn dColumnCurrent in inTable.Columns)
                    {
                        if (!IsShow(COLROW.COL, dColumnCurrent.Table.Columns.IndexOf(dColumnCurrent))) continue;
                        Write(intCurrRow, intCurrCol, dRowCurrent[dColumnCurrent].ToString());
                        intCurrCol = intCurrCol + 1 + inIntervalHoriz;
                    }

                    intCurrRow = intCurrRow + 1 + inIntervalVert;
                }
                _intLastRowWrite = intCurrRow - 1;
                _intLastColWrite = intCurrCol - 1;
                return true;
            }
            #region catch .. finally
            catch (Exception exc)
            {

                Console.WriteLine(exc.Message);
                return false;
            }
            finally
            {
                this.ColumnsNoShow.Clear();
                this.RowsNoShow.Clear();
            }
            #endregion
        }

        // Процедура проверки на нужность отображения данной строки/колонки
        private Boolean IsShow(COLROW inColRow, Int32 inIndex)
        {
            switch (inColRow)
            {
                case COLROW.COL:
                    foreach (Int32 intTemp in this.ColumnsNoShow) if (inIndex == intTemp) return false;
                    break;
                case COLROW.ROW:
                    foreach (Int32 intTemp in this.RowsNoShow) if (inIndex == intTemp) return false;
                    break;
            }
            return true;
        }

        //ЧТЕНИЕ ЗНАЧЕНИЯ ИЗ ЯЧЕЙКИ
        public string GetValue(string range)
        {
            Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            return Range.GetType().InvokeMember("Value", BindingFlags.GetProperty,
                null, Range, null).ToString();
        }

        //ЗАПИСЬ ЗНАЧЕНИЯ В ЯЧЕЙКУ
        public void SetValue(string range, string value)
        {
            Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
                null, WorkSheet, new object[] { range });
            Range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, Range, new object[] { value });
        }

        public void RunMacs(string name)
        {
            var exApp = this._excelApp;
            exApp.Visible = false;
            exApp.ScreenUpdating = false;
            exApp.Run((object) name); 
            
        }

        public void SaveFile()
        {
            var exApp = this._excelApp;
            var exBook = this.WorkBook;
            exBook.Save();
        }

        public void ExQuit()
        {
            
            this._excelApp.Quit();
        }

        public void ExQuit(bool save)
        {
            this._workBook.Close(save);
            this._excelApp.Quit();
        }
    }
}