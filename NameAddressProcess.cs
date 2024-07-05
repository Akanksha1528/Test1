using System;
using System.Collections.Generic;
using System.Linq;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace ExcelComApps
{

    public class NameAddressProcess
    {
        EXCEL.Application excelApp = null;
        EXCEL._Workbook _Workbook = null;        
        public NameAddressProcess(EXCEL.Application app, EXCEL._Workbook workbook)
        {
            excelApp = app;
            _Workbook = workbook;
        }

        private string CommentOfCellRange(EXCEL._Worksheet wsh, string colEnds, int icount)
        {
            EXCEL.Comment comment = wsh.Range[$"${colEnds}${icount}"].Comment;
            if (comment == null)
            {
                return "";
            }
            if (string.IsNullOrEmpty(comment.Text()))
            {
                return "";
            }
            return comment.Text();
        }
        public (string columnStart, int iRowStart, string columnEnd, int iRowEnd) GetDataRangeInfo(ChartFieldInfo bkInfo)
        {
            EXCEL._Worksheet wsh = null;
            var bAddress = bkInfo.SourceAddress.Replace($"={bkInfo.SheetName}!", "");
            var columnStart = bAddress.Split('$')[1];
            var rowStr = bAddress.Split('$')[2];
            rowStr = rowStr.Replace(":", "");
            int.TryParse(rowStr, out var rowStart);
            var rowEnd = rowStart;
            foreach (EXCEL._Worksheet sht in _Workbook.Worksheets)
            {
                if (sht.Name == bkInfo.SheetName)
                {
                    wsh = sht;
                    break;
                }
            }
            // Find Last Row No
            for (var icount = rowStart; icount < GlobalsPoint.UsedRows; icount++)
            {
                var questionCell = false;
                try
                {
                    var cmnt = CommentOfCellRange(wsh, columnStart, icount);
                    if (cmnt.Trim().Length > 4)
                    {
                        questionCell = true;
                    }
                    var valStr = (string)wsh.Range[$"${columnStart}${icount}"].Value.ToString();
                    if(string.IsNullOrEmpty(valStr) && !questionCell) { break; }
                    rowEnd = icount;
                }
                catch
                {
                    if (questionCell)
                    {
                        rowEnd = icount;
                        continue;
                    }
                    break;
                }
            }

            // Find Last Column Name          
            var columnEnd = columnStart;
            var colI = GeneralTask.GetColumnIndexByName(columnStart);
            for (var icount = colI; icount < GlobalsPoint.UsedColumns; icount++)
            {
                var questionCell = false;
                var cend = GeneralTask.GetColumnNameByIndex(icount);
                try
                {
                    var cmnt = CommentOfCellRange(wsh, cend, rowStart);
                    if (cmnt.Trim().Length > 4)
                    {
                        questionCell = true;
                    }
                    var valStr = (string)wsh.Range[$"${cend}${rowStart}"].Value.ToString();
                    columnEnd = cend;
                }
                catch
                {
                    if (questionCell)
                    {
                        columnEnd = cend;
                        continue;
                    }
                    break;
                }
            }
            bkInfo.ColumnStart = columnStart;
            bkInfo.ColumnEnd = columnEnd;
            bkInfo.RowStart = rowStart;
            bkInfo.RowEnd = rowEnd;
            return (columnStart, rowStart, columnEnd, rowEnd);
        }

        public List<ChartFieldInfo> CollectRangeAreaProcess(string jsonFilePath)
        {
            var existList = GeneralTask.GetChartFiledFromJson(jsonFilePath);
            GlobalsPoint.NamesInfo = existList;
            var existingChartTable = existList.Where((t => t.ContentType.ToLower() == "excel" ));
            GlobalsPoint.BookmarkNames = GetNamesInfo();
            if (!GlobalsPoint.BookmarkNames.Any())
            {
                return new List<ChartFieldInfo>();
            }
            var finalList = new List<ChartFieldInfo>();
            foreach (var eItem in existingChartTable)
            {
                var rangeInfo = GlobalsPoint.BookmarkNames.Find(r => r.BookmarkName == eItem.BookmarkName);
                if (rangeInfo == null && eItem.QuestionType.ToLower() == "contact" )
                {
                    try
                    {
                        var bkFirst = eItem.BookmarkName.Split('|');
                        rangeInfo = GlobalsPoint.BookmarkNames.Find(r => r.BookmarkName == bkFirst[1]);
                        if (rangeInfo == null)
                        {
                            continue;
                        }
                        eItem.BookmarkName = bkFirst[1];
                    }
                    catch { }
                }
                else if(rangeInfo == null)
                {
                    continue;
                }
                else
                {
                    eItem.BookmarkName = rangeInfo.BookmarkName;
                }                
                eItem.SheetName = rangeInfo.SheetName;
                eItem.SourceAddress = rangeInfo.SourceAddress == null ? "" : rangeInfo.SourceAddress;                
                finalList.Add(eItem);
            }
            return finalList;
        }
        public List<ChartFieldInfo> GetNamesInfo()
        {
            var list = new List<ChartFieldInfo>();            
            var sheetName = string.Empty;
            EXCEL.Names names = _Workbook.Names;
            foreach (EXCEL.Name name in names)
            {                
                try
                {
                    string address = (string)name.RefersTo;
                    sheetName = address.Split('!')[0].Replace("=", "");
                    string bkname = (string)name.Name;
                    if (!address.StartsWith("=#REF!"))
                    {
                        list.Add(new ChartFieldInfo { BookmarkName = bkname, SheetName = sheetName, SourceAddress = address });
                    }
                }
                catch
                {

                }
            }            
            return list;
        }

        public ChartFieldInfo GetNamesInfoByName(string bookmarkName)
        {
            foreach (EXCEL._Worksheet wsh in _Workbook.Worksheets)
            {
                var sheetName = wsh.Name;
                if (sheetName == "BL_VAR")
                {
                    continue;
                }
                EXCEL.Names names = _Workbook.Names;
                foreach (EXCEL.Name name in names)
                {
                    try
                    {
                        string address = (string)name.RefersTo;
                        string bkname = (string)name.Name;
                        if (bookmarkName == bkname)
                        {
                            if (!address.StartsWith("=#REF!"))
                            {
                                var rginfo = new ChartFieldInfo { BookmarkName = bkname, SheetName = sheetName, SourceAddress = address };
                                return rginfo;
                            }
                            return null;
                        }
                    }
                    catch
                    {

                    }
                }
            }
            return null;

        }
        public int GetColumnIndexByName(string columnName)
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
    }
}
