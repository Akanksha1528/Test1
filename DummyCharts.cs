using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using EXCEL = Microsoft.Office.Interop.Excel;
namespace ExcelComApps
{
    public class GlobalsPoint
    {
        public static List<ChartFieldInfo> BookmarkNames;

        public static List<int> BeforeAppPID { get; set; }

        public static List<int> AfterAppPID { get; set; }

        public static int HiddenCount { get; set; }

        public static EXCEL._Worksheet BLVARSHEET { get; set; }

        public static bool ChartLayoutInfoFind { get; set; }
        public static EXCEL.XlRowCol PlotedBy { get; set; }
        public static EXCEL.XlChartType ChartType { get; set; }
        public static int UsedColumns { get; set; }
        public static int UsedRows { get; set; }
        public static Dictionary<string, int> LastUsedRowsBySheet { get; set; }
        public static List<ChartColorByLegendName> _SeriesNameColorCollection { get; set; }
        public static List<ChartFieldInfo> NamesInfo { get; set; }
    }
    public class DummyCharts
    {
        public static DateTime AsseblyDateModifiedInfo()
        {
            try
            {
                string assemblyPath = Assembly.GetEntryAssembly().Location;
                FileInfo assemblyFileInfo = new FileInfo(assemblyPath);
                DateTime lastModifiedDate = assemblyFileInfo.LastWriteTime;
                return lastModifiedDate;
            }
            catch { }
            return new DateTime();
        }
        public static Version GetAssemblyVersion()
        {
            try
            {
                Assembly entryAssembly = Assembly.GetEntryAssembly();
                Version version = entryAssembly.GetName().Version;                
                return version;
            }
            catch { }
            return null;
        }
        public static int GetLastUsedRangeBySheetName(string sheetName)
        {
            try
            {
                foreach (var kvp in GlobalsPoint.LastUsedRowsBySheet)
                {
                    if (kvp.Key.ToLower() == sheetName.ToLower())
                    {
                        return kvp.Value + 3;
                    }
                }
            }
            catch { }
            return 3;
        }
        public static void SetActiveChartRange(EXCEL.Application excelApp, EXCEL.Range srg, EXCEL.XlRowCol plotBy, bool hasLegend, string imagePath, EXCEL.ChartObject chartObject)
        {
            if (!GlobalsPoint.ChartLayoutInfoFind)
            {
                Enum.TryParse(GeneralTask.GetTypeChart(excelApp, chartObject), out EXCEL.XlChartType chartTypeValue);
                GlobalsPoint.ChartType = chartTypeValue;
                GlobalsPoint.PlotedBy = chartObject.Chart.PlotBy;
            }
            try
            {
                chartObject.Activate(); 
                excelApp.ActiveChart.ChartType = GlobalsPoint.ChartType;
            }
            catch { }
            try
            {
                if (GlobalsPoint.PlotedBy == 0)
                {
                    Enum.TryParse("1", out EXCEL.XlRowCol plot);
                    GlobalsPoint.PlotedBy = plot;
                }
                if (plotBy == 0)
                {
                    excelApp.ActiveChart.SetSourceData(srg);
                }
                else
                {
                    excelApp.ActiveChart.SetSourceData(srg, GlobalsPoint.PlotedBy);
                }
                if (hasLegend)
                {
                    excelApp.ActiveChart.HasLegend = hasLegend;
                }
                if (!string.IsNullOrEmpty(imagePath))
                {
                    excelApp.ActiveChart.Export(imagePath);
                }
            }
            catch { }
        }

        public static (int rows, int columns) TotalRowsOfLegend(EXCEL.ChartObject chartObject, ref List<LegendDimentions> dimensions)
        {
            var dList = new List<double>();
            var topList = new List<double>();
            dimensions = new List<LegendDimentions>();          
            EXCEL.Legend legend = chartObject.Chart.Legend;
            var legendName = string.Empty;
            try
            {
                int findex = 1;
                int rowNo = 1; int colNo = 1;
                foreach (EXCEL.LegendEntry entry in legend.LegendEntries())
                {
                    int length = 0;
                    try
                    {
                        legendName = chartObject.Chart.FullSeriesCollection(findex).Name;
                        length = legendName.Length;
                    }
                    catch { }
                    if(!topList.Any(d => d == entry.Top) && topList.Any())
                    {
                        rowNo++; colNo++;
                    }                    
                    dList.Add(entry.Left);
                    topList.Add(entry.Top);
                    var dObj = new LegendDimentions { ColumnNo = colNo, RowNo = rowNo, Left = entry.Left, TextLength = length, Name = legendName };
                    dimensions.Add(dObj);
                    findex++;
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
            return (rows, columns);
        }

        private static void ChartAreaPercentRatio(EXCEL.ChartObject chartObject)
        {
            var percent70 = chartObject.Height * 0.7;
            var percent30 = chartObject.Height * 0.3;
            if(chartObject.Chart.Legend.Height > percent30)
            {
                chartObject.Chart.Legend.Height = percent30;
                chartObject.Chart.PlotArea.Height = percent70;
            }
            //////else if (chartObject.Chart.PlotArea.Height < percent70)
            //////{
            //////    chartObject.Chart.Legend.Height = percent70;
            //////}
            chartObject.Chart.Legend.Top = chartObject.Chart.PlotArea.Top + chartObject.Chart.PlotArea.Height;

        }
        public static void ChartAreaAdjustments(EXCEL.ChartObject chartObject, int legendCount)
        {
            var dimensions = new List<LegendDimentions>();
            if (chartObject.Chart.Legend.Position != EXCEL.XlLegendPosition.xlLegendPositionBottom)
            {
                return;
            }
            try
            {
                chartObject.Activate();
                chartObject.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                chartObject.Height = 247;
                if(chartObject.Width < 463)
                {
                    chartObject.Width = 463;
                }          
            }
            catch { }
            try
            {
                var chartHeight = chartObject.Chart.ChartArea.Height;
                var chartWidth = chartObject.Chart.ChartArea.Width;
                if (legendCount > 3)
                {
                    chartObject.Chart.Legend.Left = 0;
                    chartObject.Chart.Legend.Width = chartObject.Chart.ChartArea.Width-5;
                    chartObject.Chart.Legend.Left = 0;
                }
                chartObject.Chart.PlotArea.Top = -1;
                var plotHeight = chartObject.Chart.PlotArea.Height;
                var legendHeight = chartObject.Chart.Legend.Height;
                var rCObj = TotalRowsOfLegend(chartObject, ref dimensions);

                if (legendCount >= 5 && rCObj.rows == 1)
                {
                    chartObject.Chart.Legend.Width = chartObject.Chart.Legend.Width - 72;
                    chartObject.Chart.Legend.Left = 45;
                    legendHeight = chartObject.Chart.Legend.Height;
                    rCObj = TotalRowsOfLegend(chartObject, ref dimensions);
                }
                if(chartObject.Chart.Legend.Width < dimensions.Max(d => d.Right))
                {
                    chartObject.Chart.Legend.Top = chartObject.Chart.Legend.Top - 6;
                    chartObject.Chart.Legend.Height = chartObject.Chart.Legend.Height + 10;
                    rCObj = TotalRowsOfLegend(chartObject, ref dimensions);
                }
                if (chartObject.Chart.Legend.Width - dimensions.Max(d => d.Right) < 90 && legendCount >= 4)
                {
                    chartObject.Chart.Legend.Width = chartObject.Chart.Legend.Width - 20;
                    chartObject.Chart.Legend.Left = 10;
                    rCObj = TotalRowsOfLegend(chartObject,ref dimensions);
                }
                if (LegendOverLapCheck(ref dimensions) && legendCount >= 4)
                {
                    chartObject.Chart.Legend.Width = chartObject.Chart.Legend.Width - 20;
                    chartObject.Chart.Legend.Left = 10;
                    rCObj = TotalRowsOfLegend(chartObject, ref dimensions);
                }
                var legendReqHeight = (int)(rCObj.rows * 11.5);
               
                if(legendCount > 20)
                {
                    legendReqHeight = (int)(rCObj.rows * 11.5);
                }
                else
                {ChartAreaPercentRatio(chartObject);
                   // return;
                }                
                ChartRatioSetting(chartObject, chartHeight, legendReqHeight, legendHeight);                
            }
            catch
            {
                // Console.WriteLine($"Exception setlegend height*width {ex.Message} ");
            }
        }

        public static bool LegendOverLapCheck(ref List<LegendDimentions> dimensions)
        {
            for(int i = 0; i < dimensions.Count; i++)
            {
                int j = i + 1;
                if(j >= dimensions.Count)
                {
                    return false;
                }
                if (dimensions[j].Left - dimensions[i].Right < 90 && dimensions[i].RowNo == dimensions[j].RowNo)
                {
                    return true;
                }
            }
            return false;
        }

        private static void ChartRatioSetting(EXCEL.ChartObject chartObject, double chartHeight, int legendReqHeight, double legendHeight)
        {
            if (legendReqHeight < legendHeight)
            {
                chartObject.Chart.Legend.Height = legendReqHeight;
            }
            var Gap = chartHeight - (chartObject.Chart.Legend.Height + chartObject.Chart.Legend.Top);
            if (Gap > 3)
            {
                chartObject.Chart.Legend.Top = chartObject.Chart.Legend.Top + Gap - 1;
                var diffH = chartObject.Chart.PlotArea.Height + (chartObject.Chart.PlotArea.Top);
                chartObject.Chart.PlotArea.Height += (chartObject.Chart.Legend.Top - 2) - diffH;
                chartObject.Chart.Axes(EXCEL.XlCategoryType.xlAutomaticScale).Height = 45;
                return;
            }
            chartObject.Chart.Legend.Top = chartObject.Chart.PlotArea.Top + chartObject.Chart.PlotArea.Height;
            chartObject.Chart.Axes(EXCEL.XlCategoryType.xlAutomaticScale).Height = 45;
        }
        public static void SetLegendChart(EXCEL.ChartObject chartObject, EXCEL.XlLegendPosition lposition)
        {

            try
            {
                Microsoft.Office.Core.MsoChartElementType msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementLegendBottom;
                if (lposition == EXCEL.XlLegendPosition.xlLegendPositionTop)
                {
                    msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementLegendTop;
                }
                else if (lposition == EXCEL.XlLegendPosition.xlLegendPositionBottom)
                {
                    msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementLegendBottom;
                }
                else if (lposition == EXCEL.XlLegendPosition.xlLegendPositionRight)
                {
                    msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementLegendRight;
                }
                else if (lposition == EXCEL.XlLegendPosition.xlLegendPositionLeft)
                {
                    msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementLegendLeft;
                }
                chartObject.Chart.SetElement(msoChartElementType);
            }
            catch { }
        }
        public static void SetChartColors(EXCEL.ChartObject chartObject)
        {

            var ColorByNameCollection = GlobalsPoint._SeriesNameColorCollection;
            if (ColorByNameCollection == null)
            {
                return;
            }
            if (!ColorByNameCollection.Any())
            {
                return;
            }
            chartObject.Activate();
            EXCEL.Chart chart = chartObject.Chart;
            for (var sIndex = 1; sIndex < 200; sIndex++)
            {
                try
                {
                    var srName = chart.FullSeriesCollection(sIndex).Name;
                    if (!ColorByNameCollection.Any(n => n.Name == srName.ToString()))
                    {
                        continue;
                    }
                    var colorCode = ColorByNameCollection.Where(t => t.Name == srName.ToString()).FirstOrDefault().ColorCode;
                    if (int.TryParse(colorCode, out var code))
                    {
                        chart.FullSeriesCollection(sIndex).Interior.Color = code;
                    }
                }
                catch { }
            }
        }
        public static string CreateDummyChart(EXCEL.Application excelApp, EXCEL._Workbook wkFile, EXCEL._Worksheet wsh, EXCEL.ChartObject sourceChartObject, string imagePath, List<ChartFieldInfo> finalChartList)
        {
            var updateAddress = string.Empty;
            try
            {

                wsh.Activate();
                var title = sourceChartObject.Chart.Name;
                var objInfo = finalChartList.Find(t => t.Title == title.ToString() && t.SheetName == wsh.Name);
                try
                {
                    var rowCols = GetBookmarkAddress(excelApp, wkFile, objInfo.BookmarkName);
                    if (rowCols.rowStart == 0 || rowCols.colStart == 0) { return updateAddress; }
                    var rowColStartEnd = FillRandomValues(rowCols.rowStart, rowCols.colStart, wsh, objInfo);
                    updateAddress = ConvertRowColValue(rowColStartEnd.columnStart, rowColStartEnd.iRowStart, rowColStartEnd.columnEnd, rowColStartEnd.iRowEnd - GlobalsPoint.HiddenCount, wsh);
                }
                catch { }
                var hasLegend = sourceChartObject.Chart.HasLegend;
                var plotBy = sourceChartObject.Chart.PlotBy;
                EXCEL.Range srg = wsh.Range[updateAddress];
                sourceChartObject.Activate();
                SetActiveChartRange(excelApp, srg, plotBy, hasLegend, imagePath, sourceChartObject);
            }
            catch
            {

            }
            return updateAddress;
        }
        private static (int rowStart, int colStart) GetBookmarkAddress(EXCEL.Application excelApp, EXCEL._Workbook wkFile, string bookmarkName)
        {
            var rowNo = 0;
            var colNo = 0;
            try
            {
                var name = wkFile.Names.Item(bookmarkName);
                var address = (string)name.RefersTo;
                var splvals = address.Split('!')[1].Split('$');
                var cObj = new NameAddressProcess(excelApp, wkFile);
                colNo = cObj.GetColumnIndexByName(splvals[1]);
                int.TryParse(splvals[2], out rowNo);
            }
            catch
            {

            }
            return (rowNo, colNo);
        }
        public static (int columnStart, int iRowStart, int columnEnd, int iRowEnd) FillRandomValues(int rowNo, int ColNo, EXCEL._Worksheet wsh, ChartFieldInfo objInfo)
        {
            var actualColstart = ColNo + 1;
            var columnHeader = rowNo;
            var columnEnd = 0;

            for (var col = actualColstart; col < actualColstart + 20; col++)
            {
                try
                {
                    EXCEL.Range vCell = wsh.Cells[columnHeader, col];
                    string cellValue = vCell.Value.ToString();
                    columnEnd = col;
                }
                catch { break; }
            }
            return (ColNo, columnHeader, columnEnd, columnHeader + 1);
        }

        public static void SetRowValues(int rowNo, int colcount, int rowEnd, EXCEL._Worksheet wsh)
        {
            var preVal = "";
            var newvalStr = "";
            var colName = GeneralTask.GetColumnNameByIndex(colcount);
            for (var icount = rowNo; icount <= rowEnd; icount++)
            {
                try
                {
                    var valStr = wsh.Range[$"${colName}${icount}"].Value;
                    if (valStr != null)
                    {
                        newvalStr = (string)valStr.ToString();
                    }
                    else { newvalStr = valStr; }
                    if (preVal == newvalStr)
                    {
                        wsh.Range[$"${colName}${icount}"].Value = "";
                    }
                    preVal = newvalStr;
                }
                catch { }
            }
        }

        public static void ClubingRows(int rowNo, int colcount, int rowEnd, int colEnd, EXCEL._Worksheet wsh)
        {
            try
            {
                var colSVal = GeneralTask.GetColumnNameByIndex(colcount);
                var colEVal = GeneralTask.GetColumnNameByIndex(colEnd);
                var allRowValues = GetAllRowValues(rowNo, colcount, rowEnd, wsh);
                var reNameValues = ReNameValuseBeforeSorting(allRowValues, rowNo, rowEnd, colSVal, colEVal, wsh);
                var hiddenRowList = GetRowhasBlankValueHidden(rowNo, rowEnd, colSVal, wsh);
                GlobalsPoint.HiddenCount = hiddenRowList.Count;
                SortRowValesInAscending(rowNo, rowEnd, colSVal, colEVal, wsh);
                ShowHiddenRow(hiddenRowList, colSVal, wsh, false);
                RepaceRenamedValue(reNameValues, allRowValues, rowNo, rowEnd, colSVal, colEVal, wsh);

            }
            catch { }
        }

        public static void ShowHiddenRow(List<int> hRowList, string colS, EXCEL._Worksheet wsh, bool hide)
        {
            try
            {
                if (!hRowList.Any()) { return; }
                foreach (var hr in hRowList)
                {
                    var rng = wsh.Range[$"${colS}${hr}"];
                    rng.EntireRow.Hidden = hide;
                }
            }
            catch { }
        }
        private static void RepaceRenamedValue(List<string> reNameValues, List<ClubRowForeign> allRowValues, int rowS, int rowE, string colSVal, string colEVal, EXCEL._Worksheet wsh)
        {
            try
            {
                allRowValues.Remove(allRowValues[0]);
                var updateAddress = $"{wsh.Name}!${colSVal}${rowS}:${colEVal}${rowE}";
                for (var i = 0; i <= reNameValues.Count - 1; i++)
                {
                    var val = allRowValues.Find(e => reNameValues[i].EndsWith(e.startColValue)).startColValue;
                    var rngSeclect = wsh.Range[updateAddress];
                    rngSeclect.Replace(What: reNameValues[i], Replacement: val, LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, MatchCase: true, SearchFormat: false, ReplaceFormat: false);
                }
            }
            catch { }
        }


        private static List<string> ReNameValuseBeforeSorting(List<ClubRowForeign> allRowValues, int rowS, int rowE, string colSVal, string colEVal, EXCEL._Worksheet wsh)
        {
            var renameValues = new List<string>();
            try
            {
                var updateAddress = $"{wsh.Name}!${colSVal}${rowS}:${colEVal}${rowE}";
                var colN = 1;
                for (var i = 1; i <= allRowValues.Count - 1; i++)
                {
                    if (!string.IsNullOrEmpty(allRowValues[i].startColValue))
                    {
                        var remList = allRowValues[i];
                        var rngSeclect = wsh.Range[updateAddress];
                        var chaval = GeneralTask.GetColumnNameByIndex(colN);
                        var repVal = $"{chaval}_{allRowValues[i].startColValue}";
                        rngSeclect.Replace(What: allRowValues[i].startColValue, Replacement: repVal, LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, MatchCase: true, SearchFormat: false, ReplaceFormat: false);
                        renameValues.Add(repVal);
                        if (i != allRowValues.Count - 1)
                        {
                            i = i - 1;
                            allRowValues = allRowValues.Where(e => e.startColValue != remList.startColValue).ToList();
                        }
                    }
                    colN++;
                }
            }
            catch { }
            return renameValues;
        }

        private static void SortRowValesInAscending(int rowS, int rowE, string colSVal, string colEVal, EXCEL._Worksheet wsh)
        {
            try
            {
                var updateAddress = $"{wsh.Name}!${colSVal}${rowS}:${colEVal}${rowE}";
                EXCEL.Range rng = wsh.Range[updateAddress];
                var addressItem = $"{wsh.Name}!${colSVal}${rowS + 1}:${colSVal}${rowE}";
                EXCEL.Range itemRange = wsh.Range[addressItem];
                rng.Select();
                wsh.Sort.SortFields.Clear();
                wsh.Sort.SortFields.Add(Key: itemRange, SortOn: Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues, Order: Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, DataOption: Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);
                EXCEL.Sort dynVal = wsh.Sort;
                dynVal.SetRange(rng);
                dynVal.Header = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes;
                dynVal.MatchCase = true;
                dynVal.SortMethod = Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin;
                dynVal.Orientation = Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns;
                dynVal.Apply();
            }
            catch { }
        }

        private static List<ClubRowForeign> GetAllRowValues(int rowNo, int colcount, int rowEnd, EXCEL._Worksheet wsh)
        {
            var allRowValues = new List<ClubRowForeign>();
            var colName = GeneralTask.GetColumnNameByIndex(colcount);
            for (var icount = rowNo; icount <= rowEnd; icount++)
            {
                try
                {
                    var valStr = wsh.Range[$"${colName}${icount}"].Value;
                    if (valStr != null)
                    {
                        valStr = (string)valStr.ToString();
                    }
                    var item = new ClubRowForeign();
                    item.rowNo = icount;
                    item.startColValue = valStr;
                    allRowValues.Add(item);
                }
                catch { }

            }
            return allRowValues;
        }

        private static List<int> GetRowhasBlankValueHidden(int rowNo, int rowEnd, string colS, EXCEL._Worksheet wsh)
        {
            var rowList = new List<int>();
            for (var icount = rowNo; icount <= rowEnd; icount++)
            {
                try
                {
                    var rng = wsh.Range[$"${colS}${icount}"];
                    if (rng.Value == null)
                    {
                        continue;
                    }
                    if (rng.Value == "")
                    {
                        rng.EntireRow.Hidden = true;
                        rowList.Add(icount);
                    }
                }
                catch { }
            }
            return rowList;
        }
        private static string ConvertRowColValue(int columnStart, int iRowStart, int columnEnd, int iRowEnd, EXCEL._Worksheet wsh)
        {
            try
            {
                var colS = GeneralTask.GetColumnNameByIndex(columnStart);
                var colE = GeneralTask.GetColumnNameByIndex(columnEnd);
                return $"{wsh.Name}!${colS}${iRowStart}:${colE}${iRowEnd}";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static bool IsDuplicateInRange(EXCEL._Worksheet wsh, int rowStart, int rowEnds, int col)
        {
            var rValues = new List<string>();
            for (var rw = rowStart; rw <= rowEnds; rw++)
            {
                var fText = string.Empty;
                try
                {
                    try
                    {
                        fText = wsh.Cells[rw, col].Text;
                        if (string.IsNullOrEmpty(fText))
                        {
                            continue;
                        }
                    }
                    catch { }

                    if (rValues.Any(t => t == fText.Trim()))
                    {
                        return true;
                    }
                    else
                    {
                        rValues.Add(fText.Trim());
                    }
                }
                catch { }
            }
            return false;

        }

        public static bool SetDataLabels(EXCEL.ChartObject chartObject)
        {
            try
            {
                EXCEL.SeriesCollection seriesCollection = chartObject.Chart.SeriesCollection();
                foreach (EXCEL.Series series in seriesCollection)
                {
                    if (series.HasDataLabels)
                    {
                        EXCEL.DataLabel dataLabel = series.DataLabels(1);
                        EXCEL.XlDataLabelPosition dPostion = dataLabel.Position;
                        Microsoft.Office.Core.MsoChartElementType msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelNone;
                        if (dPostion == EXCEL.XlDataLabelPosition.xlLabelPositionCenter)
                        {
                            msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelCenter;
                        }
                        else if (dPostion == EXCEL.XlDataLabelPosition.xlLabelPositionInsideEnd)
                        {
                            msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelInsideEnd;
                        }
                        else if (dPostion == EXCEL.XlDataLabelPosition.xlLabelPositionInsideBase)
                        {
                            msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelInsideBase;
                        }
                        else if (dPostion == EXCEL.XlDataLabelPosition.xlLabelPositionOutsideEnd)
                        {
                            msoChartElementType = Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelOutSideEnd;
                        }
                        chartObject.Chart.SetElement(msoChartElementType);
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }
        public static (string bkName, bool isValid, string plotValue, int iRow, int iCol) CellNameOfRange(EXCEL._Worksheet wsh, int row, int iCol, string namePrefix)
        {
            var bkName = string.Empty;
            var iRow = 0;
            var isValid = false;
            var plotValue = "NA";
            for (iRow = row; iRow < row + 10; iRow++)
            {
                try
                {
                    bkName = wsh.Cells[iRow, iCol].Name.Name;
                    if (bkName.ToUpper().StartsWith(namePrefix.ToUpper()))
                    {
                        isValid = true;
                        break;
                    }

                    var valStr = (string)wsh.Cells[iRow, iCol].Value;
                    if (!String.IsNullOrEmpty(valStr))
                    {
                        break;
                    }

                }
                catch { }
            }
            return (bkName, isValid, plotValue, iRow, iCol);
        }
    }
}

