using SimpleOfficeCreator.Standard.Modules.Model.Component.TableDesignTab;
using System.Collections.Generic;

namespace SimpleOfficeCreator.Standard.Modules.Model.Component
{
    public class OfficeTableInfo
    {
        //public List<ColumnInfo> ColumnInfos { get; set; } = new List<ColumnInfo>();
        //public List<RowInfo> RowInfos { get; set; } = new List<RowInfo>();
        public List<OfficeModel> Children { get; set; } = new List<OfficeModel>();
        public TableCell Cell { get; set; } = new TableCell();
        public OfficeTableStyles Styles { get; set; } = new OfficeTableStyles();


        public List<int> ColumnWidthList { get; set; } = new List<int>();
        public List<int> RowHeightList { get; set; } = new List<int>();
    }

    //public class ColumnInfo
    //{
    //    public int X { get; set; }
    //    public int Width { get; set; }
    //}
    //public class RowInfo
    //{
    //    public int Y { get; set; }
    //    public int Height { get; set; }
    //}
    public class TableCell
    {
        public bool Empty { get; set; } = false;
        public int Row { get; set; } = 0;
        public int Col { get; set; } = 0;
        //세로 병합
        public int RowSpan { get; set; } = 1;

        public int ColSpan { get; set; } = 1;

        public bool VerticalMerge { get; set; } = false;

        public bool HorizontalMerge { get; set; } = false;


        /// <summary>
        /// 워드 전용
        /// </summary>
        public bool MergedRow { get; set; } = false;

        public bool IsImageCell { get; set; } = false;
    }
}
