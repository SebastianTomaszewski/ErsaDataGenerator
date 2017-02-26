using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace ErsaDataGenerator
{
    
    public class ClsGeneral
    {
        public void GenerateExcel(string pStrPath, DataTable dtSrc)
        {
            using (ExcelPackage objExcelPackage = new ExcelPackage())
            {
                
                    //Create the worksheet    
                    var objWorksheet = objExcelPackage.Workbook.Worksheets.Add(dtSrc.TableName);
                    //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1    
                    objWorksheet.Cells["A1"].LoadFromDataTable(dtSrc, true);
                    objWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));
                    objWorksheet.Cells.AutoFitColumns();
                    //Format the header    
                    using (var objRange = objWorksheet.Cells["A1:XFD1"])
                    {
                        objRange.Style.Font.Bold = true;
                        objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        objRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        objRange.Style.Fill.BackgroundColor.SetColor(Color.Blue);    
                
                }

                //Write it back to the client    
                if (File.Exists(pStrPath))
                    File.Delete(pStrPath);

                //Create excel file on physical disk    
                var objFileStrm = File.Create(pStrPath);
                objFileStrm.Close();

                //Write content to excel file    
                File.WriteAllBytes(pStrPath, objExcelPackage.GetAsByteArray());
            }
        }
    }
}
