using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using EXCEL = Microsoft.Office.Interop.Excel;
namespace ExcelComApps
{
    public class TestExcel
    {
        EXCEL.Application excelApp = null;
        EXCEL._Workbook wkFile = null;
        public void SetExcelApplication(string filePath, string sheetname)
        {
            excelApp = new EXCEL.Application()
            {
                Visible = true,
                DisplayAlerts = false
            };
            EXCEL._Workbook  wkFile = excelApp.Workbooks.Open(filePath, ReadOnly: false);
            foreach (_Worksheet wsh in wkFile.Worksheets)
            {
                if(wsh.Name != sheetname)
                {
                    continue;
                }
                EXCEL.ChartObjects sheetChartObjects = wsh.ChartObjects();
                foreach (EXCEL.ChartObject chartObject in sheetChartObjects)
                {
                    
                }
            }
        }

        public (int rows, int columns, int legendCount) TotalRowsOfLegend(EXCEL.ChartObject chartObject)
        {
            var dList = new List<double>();
            var topList = new List<double>();
            EXCEL.Legend legend = chartObject.Chart.Legend;
            try
            {
                foreach (EXCEL.LegendEntry entry in legend.LegendEntries())
                {
                    dList.Add(entry.Left);
                    topList.Add(entry.Top);
                }
            }
            catch { }
            var cp = dList.FirstOrDefault();
            var columns = dList.Distinct().Count();
            var rows = dList.Where(t => t == cp).Count();
            if (rows == 0)
            {
                rows = 1;
            }
            var legendCount = topList.Count;
            return (rows, columns, legendCount);
        }
    }
}
