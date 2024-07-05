using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using EXCEL = Microsoft.Office.Interop.Excel;
namespace ExcelComApps
{
    public class ExcelProcessApps : IDisposable
    {
        EXCEL.Application excelApp = null;
        EXCEL._Workbook wkFile = null;
        private EXCEL.XlCommentDisplayMode _commentIndicatior;
        private List<ChartFieldInfo> finalChartList = null;

        private List<CalculatedFieldDetails> calculatedFields = new List<CalculatedFieldDetails>();

        public List<ChartColorSizeDetails> ChartColorSizeDetails = new List<ChartColorSizeDetails>();
        private int docID { get; set; }
        private List<ClientChartInfo> clientChartInfos { get; set; }
        private string outJsonPath { get; set; }
        private string outCalcFieldInfoPath { get; set; }
        private int LegendCount { get; set; }
        private string BookmarkName { get; set; }
        private string NameFiledInfoJsonPath { get; set; }
        private string legendJsonPath { get; set; }
        private List<ChartLegendInfo> LegendInfoCollection { get; set; }
        public string chartImageExtension { get; set; }
        private string VersionStr { get => "Comment:Due to legend overlap, PNG format apply.Legend narrow issue if legend count > 4."; }
        private void ResetCommentIndicator(bool reset)
        {
            try
            {
                LogWrite($"Start: ResetCommentIndicator");
                if (reset)
                {
                    excelApp.DisplayCommentIndicator = _commentIndicatior;
                }
                else
                {
                    _commentIndicatior = excelApp.DisplayCommentIndicator;
                    excelApp.DisplayCommentIndicator = EXCEL.XlCommentDisplayMode.xlNoIndicator;
                }
            }
            catch { }
        }
        public string FilePath
        {
            get; set;
        }

        private void LogWrite(string msg)
        {
            try
            {
                var content = $"{DateTime.Now}     {msg}\n";
                if (File.Exists(FilePath))
                {
                    // Append the content to the existing file
                    File.AppendAllText(FilePath, content);
                }
                else
                {
                    // Create a new file and write the content to it
                    File.WriteAllText(FilePath, content);
                }
            }
            catch { }
        }
        private bool SetExcelApplication()
        {
            try
            {
                var flag = false;
#if DEBUG
                flag = true;
#endif
                if (excelApp == null)
                {
                    excelApp = new EXCEL.Application()
                    {
                        Visible = flag,
                        DisplayAlerts = false
                    };
                }
            }
            catch { }
            if (excelApp != null)
            {
                return true;
            }
            return false;
        }
        private string MakeJsonFilePath(string filepath)
        {
            var index = filepath.ToUpper().IndexOf(@"\TEMP\");
            var pre = filepath.Substring(0, index);
            var filename = Path.GetFileNameWithoutExtension(filepath);
            var fileId = Regex.Replace(filename, "[^0-9]", "");
            int.TryParse(fileId, out int id);
            docID = id;
            var dirpath = Path.Combine(pre, "temp", fileId);
            if (Directory.Exists(dirpath))
            {
                Directory.Delete(dirpath, true);
            }
            Directory.CreateDirectory(dirpath);
            return dirpath;
        }
        private List<int> GetPidByApps(string appname)
        {
            try
            {
                var ids = Process.GetProcessesByName(appname).Select(p => p.Id);
                return ids.ToList();
            }
            catch { }
            return new List<int>();
        }

        public void GetExcelSheetUsedRows(string filepath)
        {
            try
            {
                var exObj = new OpenOfficeEpPlus.OpenOffice();
                GlobalsPoint.LastUsedRowsBySheet = exObj.FetchAllSheetLastUsedRow(filepath);
            }
            catch { }
        }
        
        public void GetAllTableImageChartInformation(string filePath, string logPath)
        {

            if (!File.Exists(filePath))
            {
                return;
            }
            GetExcelSheetUsedRows(filePath);
            var resutJsonPath = MakeJsonFilePath(filePath);
            GeneralTask.NameFieldFileCopy(Path.GetDirectoryName(filePath), resutJsonPath);
            GeneralTask.ContactForeignFieldParent(Path.GetDirectoryName(filePath), resutJsonPath);
            outJsonPath = $"{resutJsonPath}\\ChartOutput.json";
            outCalcFieldInfoPath = $"{resutJsonPath}\\CalculatedFieldDetails.json";
            legendJsonPath = $"{resutJsonPath}\\LegendInfo.json";
            FilePath = Path.Combine(logPath, "chartLogs.log");
            LogWrite($"========== START : {DummyCharts.GetAssemblyVersion()} Date Modified: {DummyCharts.AsseblyDateModifiedInfo()}");
            GlobalsPoint.BeforeAppPID = GetPidByApps("excel");
            if (!SetExcelApplication())
            {
                LogWrite("Unable open excel application.");
                return;
            }
            GlobalsPoint.AfterAppPID = GetPidByApps("excel");
            try
            {
                wkFile = excelApp.Workbooks.Open(filePath, ReadOnly: false);
                NameFiledInfoJsonPath = $"{wkFile.Path}\\NameFieldInfo.json";
                if (wkFile == null)
                {
                    LogWrite("workbook not null");
                    return;
                }
            }
            catch (Exception ex)
            {
                LogWrite($"Opening workbook Exception: {ex.Message}");
                return;
            }
            try
            {
                finalChartList = CollectChartAddressArea();
            }
            catch { }
            try { RemoveDummyValueForContact(wkFile); }
            catch { }
            try { ExcelCopyForRepeatChart(); }
            catch { }
            //try { GetAllChartColorSizeDetails(wkFile, filePath);
            //    GeneralTask.AllChartColorSizeDetails(Path.GetDirectoryName(filePath), resutJsonPath);
            //}
            //catch { }
            var cInfoList = new List<ClientChartInfo>();
            cInfoList = GetAllChartInformation(wkFile, outJsonPath);
            if (cInfoList == null)
            {
                cInfoList = new List<ClientChartInfo>();
            }
            try
            {
                var tableList = GetAllTableAddress(wkFile);
                foreach (var tb in tableList)
                {
                    LogWrite($"Totale table: {tableList.Count}");
                    var cl = TableImgeInformation(tb);
                    if (cl != null)
                    {
                        cInfoList.Add(cl);
                    }
                }
                cInfoList.FirstOrDefault().docId = docID;
                cInfoList.FirstOrDefault().BeforeAppPID = GlobalsPoint.BeforeAppPID;
                cInfoList.FirstOrDefault().AfterAppPID = GlobalsPoint.AfterAppPID;
            }
            catch { }
            try
            {
                WriteJsonCalculatedField(outCalcFieldInfoPath);
            }
            catch { }
            try
            {
                WriteJsonLegendInfo(legendJsonPath);
            }
            catch { }
            clientChartInfos = cInfoList;
        }

        private void GetAllChartColorSizeDetails(EXCEL._Workbook _Workbook, string jsonPath)
        {
            try
            {
                var allColSizeDetails = new List<ChartColorSizeDetails>();
                foreach (EXCEL._Worksheet wsh in _Workbook.Worksheets)
                {
                    var shName = wsh.Name;
                    if (shName == "BL_VAR")
                    {
                        GlobalsPoint.BLVARSHEET = wsh;
                        continue;
                    }
                    try
                    {
                        wsh.Activate();
                        EXCEL.ChartObjects sheetChartObjects = wsh.ChartObjects();
                        foreach (EXCEL.ChartObject chartObject in sheetChartObjects)
                        {
                            var objInfo = finalChartList.Find(t => t.Title == chartObject.Chart.Name.ToString() && t.SheetName == wsh.Name);
                            if (objInfo == null)
                            {
                                continue;
                            }
                            var chartDetails = new ChartColorSizeDetails();
                            chartDetails.Bookmark = objInfo.BookmarkName;
                            chartDetails.ChartName = chartObject.Chart.Name;
                            chartDetails.Height = chartObject.Height.ToString();
                            chartDetails.Width = chartObject.Width.ToString();
                            var col = chartObject.Chart.ChartColor;
                            chartDetails.ColorCode = col.ToString();
                            allColSizeDetails.Add(chartDetails);
                        }
                    }
                    catch { }
                }
                if (!string.IsNullOrEmpty(jsonPath) && allColSizeDetails.Count > 0)
                {
                    var newpath = Path.Combine(Path.GetDirectoryName(jsonPath), "AllChartColorSizeDetails.json");
                    var jsonstring = JsonConvert.SerializeObject(allColSizeDetails);
                    File.WriteAllText(newpath, jsonstring);
                    ChartColorSizeDetails = allColSizeDetails;
                }
            }
            catch { }
        }
        private void ExcelCopyForRepeatChart()
        {
            var excelPath = wkFile.FullName;
            var naFile = Path.GetFileName(excelPath);
            var dirPath = Path.GetDirectoryName(outJsonPath);
            naFile = "RPT_" + naFile;
            var newFile = Path.Combine(dirPath, naFile);
            if (File.Exists(newFile))
            {
                try
                { File.Delete(newFile); }
                catch { }
            }
            File.Copy(excelPath, newFile);
        }

        public List<ChartFieldInfo> CollectChartAddressArea()
        {
            var obj = new NameAddressProcess(excelApp, wkFile);
            if (!File.Exists(NameFiledInfoJsonPath))
            {
                LogWrite("NameFieldInfo.json Not found.");
                return new List<ChartFieldInfo>();
            }
            var list = obj.CollectRangeAreaProcess(NameFiledInfoJsonPath);
            return list;
        }
        public List<ClientChartInfo> GetAllChartInformation(EXCEL._Workbook _Workbook, string chLogpath)
        {
            var allClientInfo = new List<ClientChartInfo>();
            LegendInfoCollection = new List<ChartLegendInfo>();
            LogWrite("GetAllchart...Begin");
            try
            {
                var filePath = excelApp.ActiveWorkbook.Path;
                LogWrite($"workbook FilePath  {FilePath}");
                foreach (EXCEL._Worksheet wsh in _Workbook.Worksheets)
                {
                    LogWrite("whrksheet iterations...");
                    var currentSheetName = wsh.Name;
                    if (currentSheetName == "BL_VAR")
                    {
                        GlobalsPoint.BLVARSHEET = wsh;
                        continue;
                    }
                    if (currentSheetName == "DummyColor")
                    {
                        continue;
                    }
                    wsh.Activate();
                    GlobalsPoint.UsedRows = DummyCharts.GetLastUsedRangeBySheetName(wsh.Name);
                    GlobalsPoint.UsedColumns = GlobalsPoint.UsedRows;
                    LogWrite($"Used Range sheett name: {wsh.Name} - {GlobalsPoint.UsedRows}");
                    excelApp.ActiveWorkbook.RefreshAll();
                    LogWrite("After Activate workbook not null");
                    EXCEL.ChartObjects sheetChartObjects = wsh.ChartObjects();
                    LogWrite("Sheet Get All chart objects.");
                    try
                    {
                        foreach (EXCEL.ChartObject chartObject in sheetChartObjects)
                        {
                            //LogWrite("Chart Iteration...");
                            var item = new ClientChartInfo();
                            EXCEL.Chart newChart = chartObject.Chart;
                            var chartName = newChart.Name;
                            item.Title = chartName;
                            LogWrite("before image path...");
                            var imagePath = Path.Combine(filePath, $"{item.Title}{chartImageExtension}");
                            if (File.Exists(imagePath))
                            {
                                try { File.Delete(imagePath); }
                                catch
                                {
                                }
                            }                            
                            newChart.Refresh();
                            var resp = string.Empty;
                            //LegendCount = ChartLegendCount(chartObject);
                            try { resp = EvaluteAndUpdateNewAddress(chartObject, wsh); }
                            catch { continue; }
                            if (resp == "unfilled")
                            {
#if DEBUG
#else
                                DummyCharts.CreateDummyChart(excelApp, wkFile, wsh, chartObject, imagePath, finalChartList);
                                item.FilePath = imagePath;
                                LogWrite("End Of Process...");
                                allClientInfo.Add(item);
#endif
                                continue;
                            }
                            else
                            {
                                chartObject.Activate();
                                chartObject.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
#if DEBUG
                                if (chartName == "Sheet2 Chart 5")
                                {
                                 //continue;
                                }
#endif
                                var lgObj = ChartSizeDecoration(chartObject, chartName, BookmarkName);
                                LegendInfoCollection.Add(lgObj);
                                DummyCharts.ChartAreaAdjustments(chartObject, lgObj.TotalLegend);
                                newChart.Refresh();
                                newChart.Export(imagePath);
                            }
                            item.FilePath = imagePath;
                            LogWrite("End Of Process...");
                            allClientInfo.Add(item);
                        }
                    }
                    catch (Exception exp)
                    {
                        LogWrite($"Excel sheets ChartObj iteration Exceptions {exp.Message}");
                        return allClientInfo;
                    }
                }
            }
            catch (Exception exp)
            {
                LogWrite($"EXP collection sheets charts : {exp.Message}");
            }
            return allClientInfo;
        }
        private void WriteJsonCalculatedField(string jsonPath)
        {
            try
            {
                if (calculatedFields.Count > 0)
                {
                    var jsonstring = JsonConvert.SerializeObject(calculatedFields);
                    File.WriteAllText(jsonPath, jsonstring);
                }
            }
            catch { }
        }
        private void WriteJsonLegendInfo(string jsonPath)
        {
            try
            {
                if (LegendInfoCollection.Count > 0)
                {
                    var jsonstring = JsonConvert.SerializeObject(LegendInfoCollection);
                    File.WriteAllText(jsonPath, jsonstring);
                }
            }
            catch { }
        }
        #region  "Disposing"

        private void CleanUp()
        {
            var bSucceeded = false;
            var tryCount = 0;
            var excelPath = wkFile.FullName;
            wkFile.Save();
            wkFile.Close(true, Missing.Value, Missing.Value);
            WriteJsonBeforClean(excelPath);
            Marshal.ReleaseComObject(wkFile);
            excelApp.Quit();
            if (excelApp != null)
            {
                while (!bSucceeded)
                {
                    try
                    {
                        if (tryCount > 5)
                        {

                            bSucceeded = true;
                            excelApp = null;
                            break;
                        }

                        // Cleanup:
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                        bSucceeded = true;
                    }
                    catch
                    {
                        tryCount++;
                        Thread.Sleep(500);
                    }
                }
            }
        }

        private void WriteJsonBeforClean(string FilePathExcel)
        {
            try
            {
                var base64data = GeneralTask.GetDocBlobInBase64(FilePathExcel);
                clientChartInfos.FirstOrDefault().DocBaseBlob = base64data;
                string json = JsonConvert.SerializeObject(clientChartInfos);
                File.WriteAllText(outJsonPath, json);
            }
            catch (Exception ex)
            {
                LogWrite($"Exception: {ex.Message}");
            }
        }
        private List<ChartFieldInfo> GetAllTableAddress(EXCEL._Workbook _Workbook)
        {
            var updatedTbalelist = new List<ChartFieldInfo>();
            LogWrite("start GetAllTableAddress calling.");
            foreach (EXCEL._Worksheet wsh in _Workbook.Worksheets)
            {
                var shName = wsh.Name;
                if (shName == "BL_VAR")
                {
                    GlobalsPoint.BLVARSHEET = wsh;
                    continue;
                }
                if (shName == "DummyColor")
                {
                    continue;
                }
                var listItem = EvaluateTableImage(wsh);
                updatedTbalelist.AddRange(listItem);
            }
            LogWrite("End calling GetAllTableAddress");
            return updatedTbalelist;
        }
        private ClientChartInfo TableImgeInformation(ChartFieldInfo tableFieldInfo)
        {
            LogWrite("Start TableImageIngoration");
            var sheetName = string.Empty;
            var sourceAddress = string.Empty;
            try
            {
                sheetName = tableFieldInfo.SourceAddress.Split('!')[0];
                sourceAddress = tableFieldInfo.SourceAddress.Split('!')[1];
            }
            catch { }
            if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(sourceAddress))
            {
                return null;
            }
            foreach (EXCEL._Worksheet wsh in excelApp.ActiveWorkbook.Worksheets)
            {
                if (wsh.Name.ToLower() == sheetName.ToLower())
                {
                    try
                    {
                        SortZeroValuesForTable(wsh, tableFieldInfo.RowStart, tableFieldInfo.RowEnd, tableFieldInfo.ColumnEnd);
                    }
                    catch { }
                    try
                    {
                        CollectCalulatedFieldValue(wsh);
                    }
                    catch { }
                    wsh.Activate();
                    break;
                }
            }
            ResetCommentIndicator(false);

            var ipath = Path.Combine(excelApp.ActiveWorkbook.Path, $"ExlTabImg_{DateTime.Now:mmss_fff}.png");
            try
            {
                if (File.Exists(ipath))
                {
                    File.Delete(ipath);
                }
            }
            catch { }
            var bSucceeded = false;
            var tryCount = 0;
            try
            {
                LogWrite($"Take Range");
                var worksheet = excelApp.ActiveSheet;
                EXCEL.Range dataRange = worksheet.Range[sourceAddress];
                try
                {
                    dataRange.WrapText = true;
                    dataRange.ReadingOrder = -5002;
                }
                catch { }

                LogWrite($"Start: Copy Picture");
                while (!bSucceeded)
                {
                    try
                    {
                        if (tryCount > 10)
                        {
                            bSucceeded = true;
                            break;
                        }
                        dataRange.Copy();
                        dataRange.CopyPicture(EXCEL.XlPictureAppearance.xlScreen, EXCEL.XlCopyPictureFormat.xlBitmap);
                        if (Clipboard.ContainsImage())
                        {
                            LogWrite("Containing image");
                            var image = Clipboard.GetImage();
                            image.Save(ipath);
                            var respChInfo = new ClientChartInfo { FilePath = ipath, SourceAddress = tableFieldInfo.SourceAddress, Title = tableFieldInfo.Title, Style = 606, Type = 212 };
                            LogWrite(">>> Written by CLIPBOARD...");
                            bSucceeded = true;
                            return respChInfo;
                        }
                        else
                        {
                            LogWrite(">>> Written by NO-CLIPBOARD...");
                        }
                    }
                    catch (Exception ex)
                    {
                        tryCount++;
                        Thread.Sleep(5000);
                        LogWrite($"Error:30 Dec Copy Picture: {ex.Message}");
                    }
                }
                LogWrite($"End: Copy Picture");
            }
            catch { }
            bSucceeded = false;
            tryCount = 0;
            while (!bSucceeded)
            {
                try
                {
                    if (tryCount > 10)
                    {
                        bSucceeded = true;
                        break;
                    }
                    LogWrite($"Start: Old Copy PIC");
                    var worksheet = excelApp.ActiveSheet;
                    EXCEL.Range dataRange = worksheet.Range[sourceAddress];
                    dataRange.Copy();
                    dataRange.CopyPicture(EXCEL.XlPictureAppearance.xlScreen, EXCEL.XlCopyPictureFormat.xlBitmap);
                    if (Clipboard.ContainsImage())
                    {
                        LogWrite("Containing image");
                        var image = Clipboard.GetImage();
                        image.Save(ipath);
                    }
                    var ClientChartsInfo = new ClientChartInfo { FilePath = ipath, SourceAddress = tableFieldInfo.SourceAddress, Title = tableFieldInfo.Title, Style = 606, Type = 212 };

                    EXCEL.Range sRange = worksheet.Range[sourceAddress];
                    sRange.Select();
                    try
                    {
                        sRange.WrapText = true;
                        sRange.ReadingOrder = -5002;
                        sRange = worksheet.Range[sourceAddress];
                        sRange.Select();
                        excelApp.Selection.WrapText = true;
                        excelApp.Selection.ReadingOrder = -5002;
                    }
                    catch { }
                    excelApp.Selection.CopyPicture(Appearance: EXCEL.XlPictureAppearance.xlScreen, Format: EXCEL.XlCopyPictureFormat.xlPicture);
                    EXCEL.Range prange = worksheet.Range["A1"];
                    prange.Select();
                    worksheet.Paste();
                    EXCEL.Shapes shapes = worksheet.Shapes;
                    bSucceeded = true;
                    foreach (EXCEL.Shape shape in shapes)
                    {
                        if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            continue;
                        }
                        var left = shape.Left;
                        var top = shape.Top;
                        if (left != 0 && top != 0) { continue; }
                        var width = shape.Width;
                        var height = shape.Height;
                        shape.Cut();
                        var nameList = FindAllChartName();
                        dynamic chtObj = excelApp.ActiveSheet.Shapes.AddChart2(201, EXCEL.XlChartType.xlColumnClustered);
                        var clinfo = RenderTableImage(nameList, (int)height, (int)width, tableFieldInfo);
                        LogWrite($"END: Old Copy PIC");
                        return clinfo;
                    }
                }
                catch (Exception ex)
                {
                    tryCount++;
                    Thread.Sleep(5000);
                    LogWrite($"Error:30 Dec Old Copy PIC: {ex.Message}");
                }
            }
            ResetCommentIndicator(true);
            LogWrite("End TabeImageInformation");
            return null;
        }

        private List<string> FindAllChartName()
        {
            var nameList = new List<string>();
            EXCEL.ChartObjects chartObjects = excelApp.ActiveSheet.ChartObjects();

            foreach (EXCEL.ChartObject chartObject in chartObjects)
            {
                EXCEL.Chart newChart = chartObject.Chart;
                var chartName = newChart.Name;
                nameList.Add(chartName);
            }
            return nameList;
        }
        private ClientChartInfo RenderTableImage(List<string> nameList, int height, int width, ChartFieldInfo tableInfo)
        {

            try
            {
                var filePath = excelApp.ActiveWorkbook.Path;
                EXCEL.ChartObjects chartObjects = excelApp.ActiveSheet.ChartObjects();
                foreach (EXCEL.ChartObject chartObject in chartObjects)
                {
                    EXCEL.Chart newChart = chartObject.Chart;
                    var chartName = newChart.Name;
                    if (nameList.Any(n => n == chartName))
                    {
                        continue;
                    }
                    try
                    {
                        chartObject.Chart.FullSeriesCollection(1).Delete();
                    }
                    catch { }

                    newChart.ChartArea.Height = height;
                    newChart.ChartArea.Width = width;
                    chartObject.Activate();
                    if (excelApp.ActiveChart != null)
                    {
                        excelApp.ActiveChart.Paste();
                    }
                    newChart.Refresh(); var imagetitle = tableInfo.Title;
                    var imagePath = Path.Combine(excelApp.ActiveWorkbook.Path, $"ExlTabImg_{DateTime.Now:mmss_fff}.png");
                    if (File.Exists(imagePath)) { File.Delete(imagePath); }
                    newChart.Export(imagePath);
                    var ClientChartsInfo = new ClientChartInfo { FilePath = imagePath, SourceAddress = tableInfo.SourceAddress, Title = imagetitle, Style = width, Type = height };
                    newChart.Parent.Delete();
                    return ClientChartsInfo;
                }
            }
            catch (Exception ex)
            {
                LogWrite($"Exception RenderTableImage: {ex.Message}");
            }

            return null;
        }

        private void CollectCalulatedFieldValue(EXCEL._Worksheet wsh)
        {
            try
            {
                LogWrite("Start CollectCalulatedFieldValue");
                var calList = finalChartList.Where(e => e.QuestionType == "Calculated Field");
                if (calList == null) { return; }
                foreach (var cal in calList)
                {
                    var calItem = new CalculatedFieldDetails();
                    var cObj = new NameAddressProcess(excelApp, wkFile);
                    var item = cObj.GetNamesInfoByName(cal.BookmarkName);
                    try
                    {
                        var bkRange = wsh.Range[item.SourceAddress];
                        calItem.Bookmark = cal.BookmarkName;
                        calItem.Value = bkRange.Text;
                        calItem.DocId = docID;
                        calculatedFields.Add(calItem);
                    }
                    catch { }
                }
            }
            catch { }
            LogWrite("End CollectCalulatedFieldValue");
        }
        private string EvaluteAndUpdateNewAddress(EXCEL.ChartObject chartObject, EXCEL._Worksheet wsh)
        {
            var title = chartObject.Chart.Name;            
            var objInfo = finalChartList.Find(t => t.Title == title.ToString() && t.SheetName == wsh.Name);
            if (objInfo == null)
            {
                return "unfilled";
            }            
            var cObj = new NameAddressProcess(excelApp, wkFile);
            var isOmitZeroAllow = objInfo.IsAllowZero;
            objInfo = cObj.GetNamesInfoByName(objInfo.BookmarkName);
            BookmarkName = objInfo.BookmarkName;
            if (objInfo == null)
            {
                return "unfilled";
            }
#if DEBUG
            if (objInfo.BookmarkName != "E5806672_0")
            {
                 return "unfilled";
            }
#endif
            cObj.GetDataRangeInfo(objInfo);
            var dirtyRowNames = new List<string>();
            var hideRowNames = new List<int>();
            var icount = 0;
            var rowNo = objInfo.RowStart;
            var colEnds = objInfo.ColumnEnd;
            var ColumnStart = objInfo.ColumnStart;
            var rowEnd = objInfo.RowEnd;
            var unfillRow = 0;
            var colcount = 0;
            var icolStart = cObj.GetColumnIndexByName(ColumnStart);
            var icolEnd = cObj.GetColumnIndexByName(colEnds);
            bool isNumeric = false;
            bool lessthanOne = false;
            bool isZero = false;
            GeneralTask.GetChartLayoutInformation(wsh, rowEnd, icolStart, "LF_TYPE_");

            for (icount = rowNo; icount <= rowEnd; icount++)
            {
                var rowNumericVals = 0;
                try
                {
                    for (colcount = icolStart; colcount <= icolEnd; colcount++)
                    {
                        var colName = GetColumnNameByIndex(colcount);
                        var valStr = wsh.Range[$"${colName}${icount}"].Value;
                        if (valStr != null)
                        {
                            valStr = (string)valStr.ToString();
                            IsNumberNotZero(valStr.Trim(), ref isNumeric, ref isZero, out lessthanOne);
                            if (isNumeric && !isZero)
                            {
                                rowNumericVals++;
                            }
                            if (lessthanOne)
                            {
                                wsh.Range[$"${colName}${icount}"].NumberFormat = "[$$-en-US]#,##0.00";
                            }
                        }
                    }
                }
                catch { }
                if (rowNo == rowEnd && rowNumericVals == 0)
                {
                    return "unfilled";
                }
                if (rowNumericVals == 0 && icount != rowNo)
                {
                    unfillRow = unfillRow + 1;
                }
              //  LogWrite($"isOmitZeroAllow: {isOmitZeroAllow}");
                if (isOmitZeroAllow)
                {
                    if (icount != rowNo && rowNumericVals == 0)
                    {
                        hideRowNames.Add(icount);
                        // var Dname = $"LfRw_{DateTime.Now.ToString("mmssfff")}";
                        // var refAdd = $"={wsh.Name}!${icount}:${icount}";
                        //  var rowName = wsh.Names.Add(Name: Dname, RefersTo: refAdd);
                        // dirtyRowNames.Add(icount.ToString());
                    }
                }
            }

            var hasLegend = chartObject.Chart.HasLegend;
            var plotBy = chartObject.Chart.PlotBy;
            if (rowEnd - unfillRow <= rowNo && isOmitZeroAllow)
            {
                DummyCharts.ShowHiddenRow(hideRowNames, ColumnStart, wsh, true);
                // DeleteSheetRowsByName(wsh, dirtyRowNames);
                return "unfilled";
            }
            if (isOmitZeroAllow)
            {
                DummyCharts.ShowHiddenRow(hideRowNames, ColumnStart, wsh, true);
                // DeleteSheetRowsByName(wsh, dirtyRowNames);
                rowEnd = rowEnd - dirtyRowNames.Count;
            }
            if (rowNo >= rowEnd && isOmitZeroAllow)
            {
                return "unfilled";
            }
#if DEBUG
            if (objInfo.BookmarkName == "E5808615_0")
            {
                
            }
#endif
            if (DummyCharts.IsDuplicateInRange(wsh, rowNo, rowEnd, icolStart))
            {
                DummyCharts.ClubingRows(rowNo, icolStart, rowEnd, icolEnd, wsh);
                DummyCharts.SetRowValues(rowNo, icolStart, rowEnd, wsh);
            }
            EXCEL.XlLegendPosition lPostion = EXCEL.XlLegendPosition.xlLegendPositionBottom;
            try
            {
                if (hasLegend)
                {
                    lPostion = chartObject.Chart.Legend.Position;
                }
            }
            catch { }
            var updateAddress = $"{wsh.Name}!${ColumnStart}${rowNo}:${colEnds}${rowEnd - GlobalsPoint.HiddenCount}";
            LogWrite($"Updated Address: {updateAddress}");
            EXCEL.Range srg = wsh.Range[updateAddress];
            chartObject.Activate();
            DummyCharts.SetDataLabels(chartObject);
            DummyCharts.SetActiveChartRange(excelApp, srg, plotBy, hasLegend, "", chartObject);
            LegendCount = ChartLegendCount(chartObject);
            CollectColorCodeByName(chartObject, finalChartList);
            TotalSeriesAddress(wsh, srg, hideRowNames, objInfo, "LF_TOTAL_", "", chartObject);
            DummyCharts.SetLegendChart(chartObject, lPostion);
            DummyCharts.SetChartColors(chartObject);
            return updateAddress;
        }

        private void LegendFixedNameLength(ChartObject chartObject, int displayLength)
        {
            if(LegendCount < 15) { return; }
            for (var f = 1; f < 200; f++)
            {
               
                try
                {
                    string str = chartObject.Chart.FullSeriesCollection(f).Name;
                    if (str.Length > displayLength)
                    {
                        chartObject.Chart.FullSeriesCollection(f).Name = $"{str.Substring(0, displayLength)}...";
                    }
                }
                catch { }
            }
        }
        private bool TotalSeriesAddress(EXCEL._Worksheet wsh, EXCEL.Range srg, List<int> hideRowNames, ChartFieldInfo chartFieldInfo, string namePrefix, string imagePath, ChartObject chartObject)
        {
            var isFind = false;
            var icolStart = GeneralTask.GetColumnIndexByName(chartFieldInfo.ColumnStart);
            var outInfo = DummyCharts.CellNameOfRange(wsh, chartFieldInfo.RowEnd + 1, icolStart, namePrefix);
            try
            {
                chartObject.Activate();
                excelApp.ActiveChart.ChartType = GlobalsPoint.ChartType;
                chartObject.Chart.PlotBy = GlobalsPoint.PlotedBy;
            }
            catch { }
            var response1 = DummyCharts.CellNameOfRange(wsh, chartFieldInfo.RowEnd + 1, icolStart, "LF_ORDER_");
            SortValuesDescendingOrder(wsh, chartFieldInfo.RowStart, chartFieldInfo.RowEnd, chartFieldInfo.ColumnStart, chartFieldInfo.ColumnEnd, response1.isValid, response1.bkName);
            if (!outInfo.isValid)
            {
                LegendFixedNameLength(chartObject, 22);
                return false;
            }
            try
            {
                wsh.Cells[outInfo.iRow, outInfo.iCol].value = "";
            }
            catch { }

            chartObject.Activate();
            EXCEL.Chart chart = excelApp.ActiveChart;

            // Set Stacked Column
            try
            {
                for (var f = 1; f < 200; f++)
                {
                    chart.FullSeriesCollection(f).ChartType = GlobalsPoint.ChartType;
                    if (LegendCount < 15) { continue; }
                    try
                    {
                        string str = chartObject.Chart.FullSeriesCollection(f).Name;
                        if (str.Length > 22)
                        {
                            chartObject.Chart.FullSeriesCollection(f).Name = $"{str.Substring(0, 22)}...";
                        }

                    }
                    catch { }                  
                }
            }
            catch { }

            try
            {
                // Set Legends and Remove Data Label
                chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone);
                chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelNone);
                chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementLegendBottom);
            }
            catch { }
            EXCEL.Series series = null;
            try
            {
                // Add new Series to Total for Scacked Column
                EXCEL.SeriesCollection seriesCollection = chart.SeriesCollection();
                series = seriesCollection.NewSeries();
                var nextColName = GeneralTask.GetColumnNameByIndex(icolStart + 1);
                var seriesName = $"={wsh.Name}!${chartFieldInfo.ColumnStart}${outInfo.iRow}";
                var valStr = $"={wsh.Name}!${nextColName}${outInfo.iRow}:${chartFieldInfo.ColumnEnd}${outInfo.iRow}";
                series.Name = seriesName;
                series.Values = valStr;
                series.ApplyDataLabels(Type: XlDataLabelsType.xlDataLabelsShowValue);
                series.ChartType = XlChartType.xlXYScatter;
                series.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
            }
            catch { }
            try
            {
                for (var i = 1; i <= 200; i++)
                {
                    DataLabel dataLabel = series.DataLabels(i);
                    dataLabel.Position = XlDataLabelPosition.xlLabelPositionAbove;
                    dataLabel.Font.Bold = true;
                }
            }
            catch { }
            if (!string.IsNullOrEmpty(imagePath))
            {
                excelApp.ActiveChart.Export(imagePath);
            }
            return isFind;
        }

        private void IsNumberNotZero(string str, ref bool isNumeric, ref bool isZero, out bool lessthanOne)
        {
            // bool isNumeric = Regex.IsMatch(str.Trim(), @"^[0-9]*(?:\.[0-9]*)?$");          
            // var spl = str.Replace(".", "").Trim();
            lessthanOne = false;
            var flag = double.TryParse(str, out var val);
            isNumeric = flag;
            if (val == 0)
            {
                isZero = true;
            }
            else
            {
                if (val < 1)
                {
                    lessthanOne = true; ;
                }
                isZero = false;
            }
        }

        private List<ChartFieldInfo> EvaluateTableImage(EXCEL._Worksheet wsh)
        {
            LogWrite("Start calling Evaluate Table Image");
            var namePrefix = "LF_TOTAL_";
            var objInfoList = finalChartList.Where(t => t.QuestionType == "Table Image");
            var cObj = new NameAddressProcess(excelApp, wkFile);
            var updateTableList = new List<ChartFieldInfo>();
            if (objInfoList == null)
            {
                return updateTableList;
            }
            try
            {
                foreach (var iObj in objInfoList)
                {
                    var objInfo = cObj.GetNamesInfoByName(iObj.BookmarkName);
                    if (objInfo.SourceAddress.StartsWith("=" + wsh.Name))
                    {
                        objInfo.Title = iObj.Title;
                        GlobalsPoint.UsedRows = DummyCharts.GetLastUsedRangeBySheetName(wsh.Name);
                        GlobalsPoint.UsedColumns = GlobalsPoint.UsedRows;
                        cObj.GetDataRangeInfo(objInfo);
                        var rowNo = objInfo.RowStart;
                        var colEnds = objInfo.ColumnEnd;
                        var ColumnStart = objInfo.ColumnStart;
                        var rowEnd = objInfo.RowEnd;
                        var iColStr = GeneralTask.GetColumnIndexByName(ColumnStart);
                        try
                        {
                            var outInfo = DummyCharts.CellNameOfRange(wsh, rowEnd + 1, iColStr, namePrefix);
                            if (outInfo.isValid)
                            {
                                rowEnd = outInfo.iRow;
                            }
                        }
                        catch { }
                        objInfo.SourceAddress = $"{wsh.Name}!${ColumnStart}${rowNo}:${colEnds}${rowEnd}";
                        objInfo.ColumnEnd = colEnds;
                        objInfo.ColumnStart = ColumnStart;
                        objInfo.RowEnd = rowEnd;
                        objInfo.RowStart = rowNo;
                        LogWrite($"Updated Address: {objInfo.SourceAddress}");
                        updateTableList.Add(objInfo);
                    }
                }
            }
            catch { }
            LogWrite("End calling Evaluate Table image");
            return updateTableList;
        }

        #region "Find Chart Address Range -Area"        
        private string GetColumnNameByIndex(int columnIndex)
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

        public void DeleteSheetRowsByName(EXCEL._Worksheet wsh, List<string> dirtyRowNames)
        {
            EXCEL.Names names = wkFile.Names;

            try
            {
                foreach (EXCEL.Name name in names)
                {
                    try
                    {
                        if (!dirtyRowNames.Any(t => t == name.Name || t == name.Name.Replace($"{wsh.Name}!", "")))
                        {
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        var msg = ex.Message;
                        LogWrite($"Exception Delete Row-name found: {msg}");
                    }

                    string address = (string)name.RefersTo;
                    var rw = address.Split('!')[1].Replace("$", "").Split(':')[0];
                    if (int.TryParse(rw.ToString(), out var rowToDelete))
                    {
                        wsh.Rows[rowToDelete].Delete(EXCEL.XlDeleteShiftDirection.xlShiftUp);
                        LogWrite($"Row to Delete After Delete Rows: {rowToDelete}");
                    }
                }
            }
            catch (Exception ex)
            {
                var b = ex.Message;
                LogWrite($"Exception Delete Row-2: {b}");
            }
        }

        #endregion "Find Chart Address Range -Area"

        private void RemoveDummyValueForContact(EXCEL._Workbook wkFile)
        {
            var ContactChartList = finalChartList.Where(e => e.TotalCount == 0 && e.QuestionType.ToLower() == "contact").ToList();
            if (ContactChartList == null) { return; }
            foreach (var item in ContactChartList)
            {
                EXCEL._Worksheet wsh = null;
                var cObj = new NameAddressProcess(excelApp, wkFile);
                var conList = cObj.GetNamesInfoByName(item.BookmarkName);
                var rowCol = cObj.GetDataRangeInfo(conList);
                foreach (EXCEL._Worksheet sht in wkFile.Worksheets)
                {
                    if (sht.Name == conList.SheetName)
                    {
                        wsh = sht;
                        RemoveDummyRowValue(rowCol.iRowStart, rowCol.iRowEnd, rowCol.columnStart, rowCol.columnEnd, wsh);
                        break;
                    }
                    else { continue; }
                }
            }
        }

        private void RemoveDummyRowValue(int rowS, int rowE, string colS, string colE, EXCEL._Worksheet wsh)
        {
            var colSt = GeneralTask.GetColumnIndexByName(colS);
            var colEd = GeneralTask.GetColumnIndexByName(colE);

            // First Row where Contact Created -- Write coment over Cell 
            for (var col = colSt; col <= colEd; col++)
            {
                EXCEL.Range vCell = wsh.Cells[rowS, col];
                try
                {
                    vCell.Value = vCell.Comment.Text();
                }
                catch { }
            }
            //Delete Another Row.
            for (var i = 0; i < (rowE - rowS); i++)
            {
                try
                {
                    wsh.Rows[rowS + 1].Delete(EXCEL.XlDeleteShiftDirection.xlShiftUp);
                }
                catch { }
            }
            // Insert one row
            try
            {
                EXCEL.Range rowRange = wsh.Rows[rowS + 1];
                rowRange.EntireRow.Insert(EXCEL.XlInsertShiftDirection.xlShiftDown);
            }
            catch { }
        }

        private void SortValuesDescendingOrder(EXCEL._Worksheet wsh, int rowNo, int rowEnd, string icolStart, string icolEnd, bool isOrder, string defName)
        {
            if (!isOrder) { return; }
            try
            {
                var lastInd = defName.LastIndexOf('_');
                var sortCol = defName.Substring(lastInd + 1);
                var updateAddress = $"{wsh.Name}!${icolStart}${rowNo}:${icolEnd}${rowEnd}";
                EXCEL.Range rng = wsh.Range[updateAddress];
                var addressItem = $"{wsh.Name}!${sortCol}${rowNo + 1}:${sortCol}${rowEnd}";
                EXCEL.Range itemRange = wsh.Range[addressItem];
                rng.Select();
                wsh.Sort.SortFields.Clear();
                wsh.Sort.SortFields.Add(Key: itemRange, SortOn: Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues, Order: Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending, DataOption: Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);
                EXCEL.Sort dynVal = wsh.Sort;
                dynVal.SetRange(rng);
                dynVal.Header = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes;
                dynVal.MatchCase = true;
                dynVal.SortMethod = Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin;
                dynVal.Orientation = Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns;
                dynVal.Apply();
            }
            catch { LogWrite("Error: Sorting values Descending order"); }
        }

        private void SortZeroValuesForTable(EXCEL._Worksheet wsh, int rowNo, int rowEnd, string icolEnd)
        {
            try
            {
                wsh.AutoFilterMode = false;
                LogWrite("Start Zero filter for table");
                var colEnd = GeneralTask.GetColumnIndexByName(icolEnd);
                if (!FilterRequiredForTable(wsh, rowNo, colEnd + 2, rowEnd))
                {
                    return;
                }
                var nextCol = GeneralTask.GetColumnNameByIndex(colEnd + 2);
                var addressItem = $"{wsh.Name}!${nextCol}${rowNo}:${nextCol}${rowEnd}";
                EXCEL.Range itemRange = wsh.Range[addressItem];
                itemRange.AutoFilter(1, "0");
            }
            catch { }
            LogWrite("End Zero filter for table");
        }

        private bool FilterRequiredForTable(EXCEL._Worksheet wsh, int rowStart, int ColumnNo, int rowEnd)
        {
            try
            {
                var valStr = wsh.Cells[rowStart, ColumnNo].Value;
                if (string.IsNullOrEmpty(valStr))
                {
                    return false;
                }
                if (valStr.ToString().ToLower() != "formula")
                {
                    return false;
                }

                EXCEL.Range rg = wsh.Cells[rowStart + 1, ColumnNo];
                if (rg.HasFormula)
                {
                    return true;
                }
            }
            catch { }
            return false;
        }
        private void CollectColorCodeByName(ChartObject chartObject, List<ChartFieldInfo> listInfo)
        {
            if (GlobalsPoint._SeriesNameColorCollection == null)
            {
                GlobalsPoint._SeriesNameColorCollection = new List<ChartColorByLegendName>();
            }
            var ColorByNameCollection = GlobalsPoint._SeriesNameColorCollection;

            try
            {
                chartObject.Activate();
                EXCEL.Chart chart = excelApp.ActiveChart;
                for (var sIndex = 1; sIndex < 200; sIndex++)
                {
                    try
                    {
                        var srName = chart.FullSeriesCollection(sIndex).Name;
                        string name = srName.ToString();
                        if (ColorByNameCollection.Any(n => n.Name == name))
                        {
                            continue;
                        }
                        var srColor = chart.FullSeriesCollection(sIndex).Interior.Color;
                        var newItem = new ChartColorByLegendName
                        {
                            Name = name,
                            ColorCode = srColor.ToString()
                        };

                        ColorByNameCollection.Add(newItem);
                    }
                    catch { break; }
                }
            }
            catch { }
            GlobalsPoint._SeriesNameColorCollection = ColorByNameCollection;
        }

        private ChartLegendInfo ChartSizeDecoration(EXCEL.ChartObject chartObject, string chartName, string bookmarkName)
        {
            
            var nameList = GlobalsPoint.NamesInfo;
            var wordBkName = string.Empty;
            try
            {
                if (nameList != null)
                {
                    if (nameList.Any())
                    {
                        var fobj = nameList.Where(t => t.BookmarkName == bookmarkName);
                        var qname = fobj.FirstOrDefault().QuestionName;
                        var wObj = nameList.Where(t => t.QuestionName == qname && t.ContentType.ToLower() == "word");
                        wordBkName = wObj.FirstOrDefault().BookmarkName;
                    }
                }
            }
            catch { }
            var obj = new ChartLegendInfo { Bookmark = wordBkName, Name = chartName, TotalLegend = LegendCount };
            
            if (LegendCount > 20)
                {
                    chartObject.Width = 480;
                   chartObject.Height = 247;
            }
            else
            {
                chartObject.Width = 463;
                chartObject.Height = 247;
            }            
            return obj;
        }
        private int ChartLegendCount(EXCEL.ChartObject chartObject)
        {
            var legendCount = 0;           
            try
            {
                EXCEL.Chart newChart = chartObject.Chart;
                for (var f = 1; f < 200; f++)
                {
                    var chm = newChart.FullSeriesCollection(f).ChartType;
                    legendCount++;
                }
            }
            catch { }
            LogWrite($"Legend Count: {legendCount} - {chartObject.Chart.Name}");            
            return legendCount - 1;
        }
        ~ExcelProcessApps()
        {

        }
        public void Dispose()
        {
            Dispose(true);
            // Now since we have done the cleanup already there is nothing left
            // for the Finalizer to do. So lets tell the GC not to call it later.
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            CleanUp();
            if (disposing == true)
            {
                //someone want the deterministic release of all resources
                //Let us release all the managed resources
            }
            else
            {
                // Do nothing, no one asked a dispose, the object went out of
                // scope and finalized is called so lets next round of GC 
                // release these resources
            }

            // Release the unmanaged resource in any case as they will not be 
            // released by GC

        }
        #endregion "Disposing"
    }
}
