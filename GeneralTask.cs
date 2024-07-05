using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace ExcelComApps
{
    static class GeneralTask
    {
        public static bool NameFieldFileCopy(string sourcePath, string destPath)
        {
            try
            {
                var filepath = Path.Combine(sourcePath, "NameFieldInfo.json");
                if (!File.Exists(filepath))
                {
                    return false;
                }
                File.Copy(filepath, Path.Combine(destPath, "NameFieldInfo.json"));
                return true;
            }
            catch { }
            return false;
        }
        public static bool ContactForeignFieldParent(string sourcePath, string destPath)
        {
            try
            {
                var filepath = Path.Combine(sourcePath, "ContactForeignFieldParent.json");
                if (!File.Exists(filepath))
                {
                    return false;
                }
                File.Copy(filepath, Path.Combine(destPath, "ContactForeignFieldParent.json"));
                return true;
            }
            catch { }
            return false;
        }
        public static bool AllChartColorSizeDetails(string sourcePath, string destPath)
        {
            try
            {
                var filepath = Path.Combine(sourcePath, "AllChartColorSizeDetails.json");
                if (!File.Exists(filepath))
                {
                    return false;
                }
                File.Copy(filepath, Path.Combine(destPath, "AllChartColorSizeDetails.json"));
                return true;
            }
            catch { }
            return false;
        }
        public static int GetColumnIndexByName(string columnName)
        {
            var columnIndex = 0;
            var position = 0;

            try
            {
                for (var i = columnName.Length - 1; i >= 0; i--)
                {
                    var charValue = columnName[i] - 'A' + 1;
                    columnIndex += charValue * (int)Math.Pow(26, position);
                    position++;
                }
            }
            catch { }
            return columnIndex;
        }
        public static string GetColumnNameByIndex(int columnIndex)
        {
            var columnName = "";
            try
            {

                while (columnIndex > 0)
                {
                    var remainder = (columnIndex - 1) % 26;
                    var charValue = (char)('A' + remainder);
                    columnName = charValue + columnName;
                    columnIndex = (columnIndex - 1) / 26;
                }
            }
            catch { }
            return columnName;
        }

        public static List<ChartFieldInfo> GetChartFiledFromJson(string filepath)
        {
            try
            {
                var txt = File.ReadAllText(filepath);
                var obj = JsonConvert.DeserializeObject<List<ChartFieldInfo>>(txt);
                return obj;
            }
            catch { }
            return new List<ChartFieldInfo>();
        }

        public static string GetDocBlobInBase64(string filePath)
        {
            if (File.Exists(filePath))
            {
                var blob = File.ReadAllBytes(filePath);
                return Convert.ToBase64String(blob);
            }
            return null;
        }

        public static void GetChartLayoutInformation(EXCEL._Worksheet wsh, int irowEnd, int icolStart, string namePrefix)
        {
            var outInfo = DummyCharts.CellNameOfRange(wsh, irowEnd + 1, icolStart, namePrefix);

            if (!outInfo.isValid)
            {
                outInfo = DummyCharts.CellNameOfRange(wsh, irowEnd + 1, icolStart + 1, namePrefix);
                if (!outInfo.isValid)
                {
                    GlobalsPoint.ChartLayoutInfoFind = false;
                    return;
                }
            }
            var flagName = string.Empty;
            try
            {
                flagName = outInfo.bkName.Substring(0, outInfo.bkName.LastIndexOf("_"));
            }
            catch { }
            GlobalsPoint.ChartLayoutInfoFind = true;

            switch (flagName)
            {
                case "LF_TYPE_CLUSTER_ROW":
                    GlobalsPoint.PlotedBy = EXCEL.XlRowCol.xlRows;
                    GlobalsPoint.ChartType = EXCEL.XlChartType.xlColumnClustered;
                    break;
                case "LF_TYPE_CLUSTER_COL":
                    GlobalsPoint.PlotedBy = EXCEL.XlRowCol.xlColumns;
                    GlobalsPoint.ChartType = EXCEL.XlChartType.xlColumnClustered;
                    break;
                case "LF_TYPE_STACK_COL":
                    GlobalsPoint.PlotedBy = EXCEL.XlRowCol.xlColumns;
                    GlobalsPoint.ChartType = EXCEL.XlChartType.xlColumnStacked;
                    break;
                case "LF_TYPE_STACK_ROW":
                    GlobalsPoint.PlotedBy = EXCEL.XlRowCol.xlRows;
                    GlobalsPoint.ChartType = EXCEL.XlChartType.xlColumnStacked;
                    break;
                default:
                    GlobalsPoint.ChartLayoutInfoFind = false;
                    break;
            }
        }

        public static (string bkName, bool isValid, string plotValue, int iRow, int iCol) CellNameOfRange(EXCEL._Worksheet wsh, int row, int iCol, string namePrefix, string plotPrefix)
        {
            var bkName = string.Empty;
            var iRow = 0;
            var isValid = false;
            var plotValue = "NA";
            for (iRow = row; iRow < row + 5; iRow++)
            {
                try
                {
                    bkName = wsh.Cells[iRow, iCol].Name.Name;
                    if (bkName.ToUpper().StartsWith(namePrefix.ToUpper()))
                    {
                        isValid = true;
                    } // plotPrefix
                    else if (bkName.ToUpper().StartsWith(plotPrefix.ToUpper()))
                    {
                        plotValue = bkName;
                    }
                    if (isValid == true && !string.IsNullOrEmpty(plotValue))
                    {
                        break;
                    }

                }
                catch { }
            }
            return (bkName, isValid, plotValue, iRow, iCol);
        }

        public static string GetTypeChart(EXCEL.Application excelApp, EXCEL.ChartObject chartObject)
        {
            try
            {
                var valueToFind = chartObject.Chart.Name;
                EXCEL._Worksheet excelWorksheet = GlobalsPoint.BLVARSHEET;
                EXCEL.Range searchRange = excelWorksheet.UsedRange;
                EXCEL.Range foundCell = searchRange.Find(valueToFind, Type.Missing, EXCEL.XlFindLookIn.xlValues, EXCEL.XlLookAt.xlPart);
                if (foundCell != null)
                {
                    try
                    {
                        var colrVal = excelWorksheet.Cells[foundCell.Row, foundCell.Column + 2].Value;
                        return colrVal;
                    }
                    catch { };
                    try
                    {
                        var colrVal1 = excelWorksheet.Cells[foundCell.Row, foundCell.Column + 2].Formula;
                        return colrVal1;
                    }
                    catch { };
                }
            }
            catch { }
            return string.Empty;
        }


    }
}
