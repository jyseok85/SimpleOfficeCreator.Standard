using SimpleOfficeCreator.Stardard.Modules.Model.Component;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.PictureFormatTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.ShapeFormat;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.TableDesignTab;
using System;
using System.Collections.Generic;

namespace SimpleOfficeCreator.Stardard.Modules.Model
{
    /// <summary>
    /// 컨트롤 생성 순서에 따라서 컨트롤의 z-index 가 할당된다.
    /// </summary>
    public class OfficeModelCreator
    {
        //파워포인트와 
        //워드의 테이블구조가 개판이다.

        //파워포인트는 원본셀 속성만 바라본다. 그러나 모든 빈셀을 만들어둬야 한다. 
        //워드는 각 개별 셀 속성을 바라본다. 그러나 워드는 개별 셀들을 전부 만들지 않고, 세로 병합되는 셀만 생성한다. 가로 병합 빈셀은 안만든다.

        public OfficeModel CreateTextBox(int x, int y, int width, int height, string text, OfficeFont font = null, OfficeParagraph paragraph = null, OfficeShapeStyle style = null)
        {
            OfficeModel model = new OfficeModel("");
            if (font is null)
                model.Font = new OfficeFont();
            else
                model.Font = font;
            if (paragraph is null)
                model.Paragraph = new OfficeParagraph();
            else
                model.Paragraph = paragraph;
            if (style is null)
                model.ShapeStyle = new OfficeShapeStyle();
            else
                model.ShapeStyle = style;

            model.Rect.X = x;
            model.Rect.Y = y;
            model.Rect.Width = width;
            model.Rect.Height = height;
            model.Text = text;
            model.Type = Type.TextBox;
            return model;
        }

        /// <summary>
        /// !셀의 내용이 Width 보다 커질경우 WordWrap 속성적용으로 인하여 Row Height가 늘어나게 된다(비활성 불가능). Column Width 설정에 주의하자.
        /// </summary>
        public OfficeModel CreateTable(int x, int y, int width, int height, List<int> colWidths, List<int> rowHeights)
        {
            //테이블 생성할 차례.. 잘되고있어요 굿굿   
            OfficeModel model = new OfficeModel("");
            model.Rect.X = x;
            model.Rect.Y = y;
            model.Rect.Width = width;
            model.Rect.Height = height;
            model.Type = Type.Table;
            model.TableInfo = new OfficeTableInfo();
            model.TableInfo.ColumnWidthList = colWidths;
            model.TableInfo.RowHeightList = rowHeights;
            return model;
        }

        /// <summary>
        /// 테이블 셀의 경우 순서대로 생성해야 합니다. 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="row">1부터 시작</param>
        /// <param name="col">1부터 시작</param>
        /// <param name="text"></param>
        /// <param name="rowSpan"></param>
        /// <param name="colSpan"></param>
        /// <param name="font"></param>
        /// <param name="paragraph"></param>
        /// <param name="style"></param>
        public void CreateTableCell(OfficeModel parent, int row, int col, string text, int rowSpan, int colSpan, OfficeFont font = null, OfficeParagraph paragraph = null, OfficeTableStyles style = null)
        {
            if (parent.TableInfo is null)
            {
                Logger.Instance.Write("TableInfo가 없습니다.");
                return;
            }

            //if (parent.TableInfo.ColumnWidthList.Count < col)
            //{
            //    int cellWidthSum = parent.TableInfo.ColumnWidthList.Sum() + colWidth;
            //    if (parent.Rect.Width < cellWidthSum)
            //        throw new Exception(Logger.Instance.Write("셀 Width의 총합이 테이블의 Width보다 클수는 없습니다."));
            //    else
            //        parent.TableInfo.ColumnWidth.Add(colWidth);
            //}

            //if (parent.TableInfo.RowHeight.Count < row)
            //{
            //    int cellHeightSum = parent.TableInfo.RowHeight.Sum() + rowHeight;
            //    if (parent.Rect.Height < cellHeightSum)
            //        throw new Exception(Logger.Instance.Write("셀 Height의 총합이 테이블의 Height보다 클수는 없습니다."));
            //    else
            //        parent.TableInfo.RowHeight.Add(rowHeight);
            //}

            OfficeModel model = new OfficeModel("");
            model.Type = Type.TableCell;
            model.TableInfo = new OfficeTableInfo();
            model.TableInfo.Cell.Row = row;
            model.TableInfo.Cell.Col = col;
            model.TableInfo.Cell.RowSpan = rowSpan;
            model.TableInfo.Cell.ColSpan = colSpan;
            model.Text = text;

            if (font is null)
                model.Font = new OfficeFont();
            else
                model.Font = font;
            if (paragraph is null)
                model.Paragraph = new OfficeParagraph();
            else
                model.Paragraph = paragraph;
            if (style is null)
                model.TableInfo.Styles = new OfficeTableStyles();
            else
                model.TableInfo.Styles = style;

            parent.TableInfo.Children.Add(model);



            //셀 내용넣고
            if (rowSpan > 1 || colSpan > 1)
            {
                bool verticlaMerge = rowSpan > 1 ? true : false;

                for (int i = 0; i < rowSpan; i++)
                {
                    CreateEmptyCell(col, row, verticlaMerge, 0, i, colSpan);
                }
                for (int j = 0; j < colSpan; j++)
                {
                    CreateEmptyCell(col, row, verticlaMerge, j, 0, 1);
                }
            }


            void CreateEmptyCell(int c, int r, bool isRowSpan, int colindex, int rowindex, int colspan)
            {
                OfficeModel empty = new OfficeModel("emptycell");
                empty.TableInfo = new OfficeTableInfo();
                empty.TableInfo.Cell.Row = r + rowindex;
                empty.TableInfo.Cell.Col = c + colindex;
                empty.TableInfo.Cell.Empty = true;


                empty.TableInfo.Cell.HorizontalMerge = true;
                empty.TableInfo.Cell.VerticalMerge = isRowSpan;

                //워드 전용속성
                if (colspan > 1)
                    empty.TableInfo.Cell.ColSpan = colspan;

                if (rowindex > 0)
                {
                    empty.TableInfo.Cell.MergedRow = true;
                    empty.TableInfo.Styles = style;
                }

                parent.TableInfo.Children.Add(empty);
            }
        }

        public List<OfficeModel> End(OfficeModel model)
        {
            List<OfficeModel> listTableCell = new List<OfficeModel>();
            if (model.Type == Type.Table)
            {
                var xList = model.TableInfo.ColumnWidthList;
                var yList = model.TableInfo.RowHeightList;

                for (int i = 0; i < xList.Count; i++)
                {
                    for (int j = 0; j < yList.Count; j++)
                    {
                        var result = model.TableInfo.Children.Find(x => x.TableInfo.Cell.Col == i && x.TableInfo.Cell.Row == j);
                        if (result != null)
                        {
                            listTableCell.Add(result);
                        }
                        else
                        {
                            //병합을 위한 빈 모델을 만든다.
                            OfficeModel empty = new OfficeModel("emptycell");
                            empty.TableInfo = new OfficeTableInfo();
                            empty.TableInfo.Cell.Row = j;
                            empty.TableInfo.Cell.Col = i;
                            empty.TableInfo.Cell.Empty = true;

                            empty.TableInfo.Cell.HorizontalMerge = true;
                            empty.TableInfo.Cell.VerticalMerge = false;
                            listTableCell.Add(empty);
                        }
                    }
                }
            }
            return listTableCell;
        }

        public OfficeModel CreatePicture(int x, int y, int width, int height, string base64, OfficePictureStyle style = null)
        {
            OfficeModel model = new OfficeModel("");
            if (style is null)
                model.PictureStyle = new OfficePictureStyle();
            else
                model.PictureStyle = style;

            model.Rect.X = x;
            model.Rect.Y = y;
            model.Rect.Width = width;
            model.Rect.Height = height;
            model.Type = Type.Picture;
            model.Text = base64;
            var ticks = new DateTime(2016, 1, 1).Ticks;
            var ans = DateTime.Now.Ticks - ticks;
            var uniqueId = ans.ToString("x");

            model.UID = "id_" + uniqueId;
            return model;
        }

        public OfficeModel CreateShape(int x, int y, int width, int height, OfficeShapeStyle style = null)
        {
            OfficeModel model = new OfficeModel("");
            if (style is null)
                model.ShapeStyle = new OfficeShapeStyle();
            else
                model.ShapeStyle = style;

            model.Rect.X = x;
            model.Rect.Y = y;
            model.Rect.Width = width;
            model.Rect.Height = height;
            model.Type = Type.Shape;
            return model;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="width">용지 가로 크기</param>
        /// <param name="height">용지 세로 크기</param>
        /// <param name="landscape">가로 여부</param>
        /// <returns></returns>
        public OfficeModel CreateReport(int width, int height, bool landscape, float marginLeft, float marginTop, float marginRight, float marginBottom)
        {
            OfficeModel model = new OfficeModel("");
            model.PaperInfo = new PaperInfo();
            model.Margin.Left = marginLeft;
            model.Margin.Top = marginTop;
            model.Margin.Right = marginRight / 2; //왜인지 DR에 두배로 들어가있다..
            model.Margin.Bottom = marginBottom / 2;
            model.PaperInfo.Width = width;
            model.PaperInfo.Height = height;
            model.PaperInfo.IsLandscape = landscape;
            model.Type = Type.Paper;
            return model;
        }
    }
}
