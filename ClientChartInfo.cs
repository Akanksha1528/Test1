using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelComApps
{
    public class ClientChartInfo
    {
        public string FilePath { get; set; }
        public string Title { get; set; }
        public int Style { get; set; }
        public int Type { get; set; }
        public string SourceAddress { get; set; }
        public List<int> BeforeAppPID { get; set; }
        public List<int> AfterAppPID { get; set; }
        public int docId { get; set; }
        public string DocBaseBlob { get; set; }
    }
    public class ChartFieldInfo
    {
        public string BookmarkName { get; set; }
        public string Title { get; set; }
        public int Style { get; set; }
        public int Type { get; set; }
        public string SheetName { get; set; }
        public string SourceAddress { get; set; }
        public string ColumnStart { get; set; }
        public string ColumnEnd { get; set; }
        public int RowStart { get; set; }
        public int RowEnd { get; set; }
        public string QuestionName { get; set; }
        public string QuestionType { get; set; }
        public string ContentType { get; set; }
        public int TotalCount { get; set; }

        /// <summary>
        /// if true: zero related bar/column will not appear
        /// </summary>
        public bool IsAllowZero { get; set; }
    }

    public class ClubRowForeign
    {
        public int rowNo { get; set; }
        public string startColValue { get; set; }
    }
    
    public class ChartColorSizeDetails
    {
        public string ChartName { get; set; }
        public string ColorCode { get; set; }
        public string Bookmark { get; set; }
        public string Height { get; set; }
        public string Width { get; set; }
    }
    public class CalculatedFieldDetails
    {
        public string Bookmark { get; set; }
        public string Value { get; set; }
        public int DocId { get; set; }
    }
    public class ChartColorByLegendName
    {
        public string Name { get; set; }
        public string ColorCode { get; set; }       
    }

    public class ChartLegendInfo
    {
        public string Name { get; set; }
        public string Bookmark { get; set; }
        public int TotalLegend { get; set; }
    }

    public class LegendDimentions
    {
        public int RowNo { get; set; }
        public int ColumnNo { get; set; }
        public string Name { get; set; }
        public double TextLength { get; set; }
        public double Left { get; set; }
        public double Right { get => Left + TextLength; }
    }
}
