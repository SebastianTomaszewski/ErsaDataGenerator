using System;
using System.Data;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static System.String;

namespace ErsaDataGenerator
{
    public class Excel
    {
        private string _filePath;
        private string _fileName;

        private string FilePath
        {
            get
            {
                return !IsNullOrEmpty(_filePath)
                    ? _filePath
                    : (_filePath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        $@"{FileName}_{DateTime.Today.ToShortDateString()}.xlsx"));
            }
            set { _filePath = value; }
        }
        private string FileName
        {
            get { return !IsNullOrEmpty(_fileName) ? _fileName : _fileName = "Data"; }
            set { _fileName = value; }
        }


        #region public metod

        public void GenerateExcel(DataTable dtSrc, string pStrPath = "",string name ="")
        {
            FileName = name;
            FilePath = pStrPath;
          
            using (var objPackage = new ExcelPackage())
            {
                //Create WorkSheet
                var objWorksheet = CreateSheet(objPackage, FileName);

                //Add data
                AddDataTable(objWorksheet, "A1", dtSrc);

                //Add header style
                AddStyleHeader(objWorksheet, 1, 1, dtSrc.Columns.Count, 1);

                //save file
                Save(objPackage, FilePath);
            }
        }

        #endregion

        #region private metod

        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            var ws = p.Workbook.Worksheets[1];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }

        private static void Save(ExcelPackage p, string path)
        {
            //Write it back to the client    
            if (File.Exists(path))
                File.Delete(path);

            //Create excel file on physical disk    
            var objFileStrm = File.Create(path);
            objFileStrm.Close();

            //Write content to excel file    
            File.WriteAllBytes(path, p.GetAsByteArray());
        }

        private static void AddStyleHeader(ExcelWorksheet ws, int fromCol, int fromRow, int toCol, int toRow)
        {
            using (var objRange = ws.Cells[fromRow, fromCol, toRow, toCol])
            {
                objRange.Style.Font.Bold = true;
                objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                objRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                objRange.Style.Fill.BackgroundColor.SetColor(Color.Black);
                objRange.Style.Font.Color.SetColor(Color.White);
            }
        }

        private static void AddDataTable(ExcelWorksheet ws, string cell, DataTable dt)
        {
            ws.Cells[cell].LoadFromDataTable(dt, true);
        }

        #endregion

        #region Comment - do rozbudowy w dll

        ////public class Excel : IDisposable
        //#region ctr
        //public Excel(bool isManual)
        //{
        //    _objExcelPackage = new ExcelPackage();
        //}
        // #endregion
        //#region file

        //private ExcelPackage _objExcelPackage;
        //private ExcelWorksheet _objExcelWorksheet;

        //private string _fileName;
        //private string _folderPath;

        //#endregion
        //#region propertis

        //public string FileName
        //{
        //    get
        //    {
        //        return !IsNullOrEmpty(_fileName) ? $@"{_fileName}.xlsx" : $@"{DateTime.Today.ToShortDateString()}.xlsx";
        //    }
        //    set
        //    {
        //        value = value.Replace('<', '_').Replace('>', '_').Replace('*', '_');
        //        value = value.Replace('/', '_').Replace('?', '_').Replace('|', '_');
        //        value = value.Replace(';', '_').Replace(':', '_').Replace('"', '_');
        //        _fileName = value;
        //    }
        //}

        //public string FilePath => Path.Combine(FolderPath, FileName);

        //public string FolderPath
        //{
        //    get
        //    {
        //        return !IsNullOrEmpty(_folderPath)
        //            ? FilePath
        //            : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //    }
        //    set { _folderPath = value; }
        //}

        //#endregion
        //#region Public metod

        //public void AddDataTable(string cell, DataTable dt)
        //{
        //    // if objExcelWorksheet is null - create
        //    _objExcelWorksheet.Cells[cell].LoadFromDataTable(dt, true);
        //}
        //public void CreateSheet(string sheetName)
        //{
        //    _objExcelPackage.Workbook.Worksheets.Add(sheetName);
        //    _objExcelWorksheet = _objExcelPackage.Workbook.Worksheets[1];
        //    _objExcelWorksheet.Name = sheetName; //Setting Sheet's name
        //    _objExcelWorksheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
        //    _objExcelWorksheet.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet         
        //}

        //#endregion
        //public void Dispose()
        //{
        //    _objExcelPackage?.Dispose();
        //}

        #endregion
    }
}
