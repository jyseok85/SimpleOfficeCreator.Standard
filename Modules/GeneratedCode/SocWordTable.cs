using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.Model;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.TableDesignTab;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class SocWordTable
    {

        private SocWordTable() { }
        //private static 인스턴스 객체
        private static readonly Lazy<SocWordTable> _instance = new Lazy<SocWordTable>(() => new SocWordTable());
        //public static 의 객체반환 함수
        public static SocWordTable Instance { get { return _instance.Value; } }

        const int wordTableRatio = 15;

        public Table Generate(OfficeModel model)
        {
            bool isSingleTable = model.TableInfo is null ? true : false;
            var table = new Table();

            PointF loc = new PointF((model.Rect.X) * wordTableRatio, (model.Rect.Y) * wordTableRatio);
            table.Append(GenerateTableProperties(!isSingleTable, loc));

            if (isSingleTable)
                table.Append(GenerateTableGrid(new List<int> { model.Rect.Width }));
            else
                table.Append(GenerateTableGrid(model.TableInfo.ColumnWidthList));

            if (isSingleTable)
            {
                TableRow tRow = GenerateTableRow(model, model.Rect.Height, -1);
                table.Append(tRow);
            }
            else
            {
                for (int i = 0; i < model.TableInfo.RowHeightList.Count; i++)
                {
                    var height = model.TableInfo.RowHeightList[i];
                    TableRow tRow = GenerateTableRow(model, height, i);
                    table.Append(tRow);
                }
            }
            //tableHeight = table.Height;

            return table;
        }


        public TableProperties GenerateTableProperties(bool isTableLabel, PointF loc)
        {
            //Y 값이 0 일경우에 OpenXml 내부적으로 기존테이블과 합쳐버리는 로직이 있다...(아마도..)
            //그래서 어차피 눈에도 안보이는데 1증가시킨다. 실제 테이블할때는 제외하도록하자.
            if (loc.Y == 0 && isTableLabel == false)
            {
                loc.Y = 1;
            }

            TableProperties tableProperties1 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties()
            {
                LeftFromText = 142,
                RightFromText = 142,
                VerticalAnchor = VerticalAnchorValues.Text,
                HorizontalAnchor = HorizontalAnchorValues.Text,
                TablePositionX = (int)loc.X,
                TablePositionY = (int)loc.Y
            };

            //TableOverlap tableOverlap1 = new TableOverlap(){ Val = TableOverlapValues.Overlap };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Dxa }; //자동크기 조정 X
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };

            //전혀 영향이 없는듯한데... TableCellMargin을 설정하면 무시된다.
            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa };
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            TableLook tableLook1 = new TableLook() { Val = "0000", FirstRow = false, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tablePositionProperties1);
            //tableProperties1.Append(tableOverlap1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);
            return tableProperties1;
        }

        public TableGrid GenerateTableGrid(List<int> columnWidths)
        {
            TableGrid tableGrid = new TableGrid();

            foreach (int colWidth in columnWidths)
            {
                var width = (colWidth ) * wordTableRatio;
                GridColumn gridColumn1 = new GridColumn() { Width = width.ToString() };
                tableGrid.Append(gridColumn1);
            }

            return tableGrid;

        }

        public TableRow GenerateTableRow(OfficeModel model, int height, int row)
        {
            var tableRow = new TableRow();
            TableRowProperties tableRowProperties1 = new TableRowProperties();
            tableRowProperties1.Append(new TableRowHeight() { Val = (uint)(height * wordTableRatio), HeightType = HeightRuleValues.Exact });
            tableRow.Append(tableRowProperties1);

            if (row == -1)
            {
                //todo : 단일 셀일경우 
                var tableCell = GenerateTableCell(model, true);
                tableRow.Append(tableCell);
            }
            else
            {
                if (model.TableInfo.Children != null)
                {
                    List<OfficeModel> items = model.TableInfo.Children.FindAll(x => x.TableInfo.Cell.Row == row);
                    foreach (var item in items)
                    {
                        if (item.TableInfo.Cell.Empty == true)
                        {
                            if(item.TableInfo.Cell.MergedRow == true) 
                            {
                                var tableCell = GenerateEmptyTableCell(item);
                                tableRow.Append(tableCell);
                            }
                        }
                        else
                        {
                            var tableCell = GenerateTableCell(item);
                            tableRow.Append(tableCell);
                        }
                    }
                }
            }
            return tableRow;
        }

        TableCell GenerateTableCell(OfficeModel cell, bool isSingleCell = false)
        {
            TableCell tableCell1 = new TableCell();

            //1. 텍스트 박스 설정
            tableCell1.Append(SetTextBody());

            //2. 테이블 셀 모양설정 : 텍스트 방향, 수직 정렬, 배경색, 테두리 모양 등
            tableCell1.Append(SetCellProperty());

            return tableCell1;

            Paragraph SetTextBody()
            {
                //단락 : 단락 속성과 텍스트로 구성됨.
                Paragraph paragraph1 = new Paragraph();
                var paragraphProperties = new ParagraphProperties();
                #region 공백없음 속성
                paragraphProperties.Append(new ParagraphStyleId() { Val = "a3" });
                #endregion

                #region 가로정렬
                paragraphProperties.Append(new Justification() { Val = Common.Instance.GetWordprocessingJustification(cell.Paragraph.AlignmentHorizontal) });
                #endregion

                #region 줄간격
                if (cell.Paragraph.LineSpacing > 0)
                {
                    var lineSpace = cell.Paragraph.LineSpacing * 20f;
                    SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = lineSpace.ToString(), LineRule = LineSpacingRuleValues.Exact };
                    paragraphProperties.Append(spacingBetweenLines1);
                }
                #endregion

                paragraph1.Append(paragraphProperties);

                //2. 텍스트 속성 : 폰트, 컬러, 내용, Bold 등
                paragraph1.Append(SetRun());

                //todo : 이미지 셀 확인

                return paragraph1;

                Run SetRun()
                {
                    var run = new Run();
                    var runProperties1 = Common.Instance.GetWordRunProperty(cell.Font);

                    //#region 폰트명
                    //RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = cell.Font.Name, HighAnsi = cell.Font.Name, EastAsia = cell.Font.Name };
                    //runProperties1.Append(runFonts4);
                    //#endregion

                    //#region 폰트사이즈
                    //FontSize fontSize4 = new FontSize() { Val = (cell.Font.Size * 2).ToString() };
                    //runProperties1.Append(fontSize4);
                    //#endregion

                    //#region 폰트옵션
                    //if (cell.Font.UnderLine)
                    //{
                    //    Underline style = new Underline() { Val = UnderlineValues.Single };
                    //    runProperties1.Append(style);
                    //}
                    //if (cell.Font.Strike)
                    //{
                    //    Strike style = new Strike();
                    //    runProperties1.Append(style);
                    //}
                    //if (cell.Font.Bold)
                    //{
                    //    Bold style = new Bold();
                    //    runProperties1.Append(style);
                    //}
                    //if (cell.Font.Italic)
                    //{
                    //    Italic style = new Italic();
                    //    runProperties1.Append(style);
                    //}
                    //#endregion

                    //#region 폰트컬러
                    //runProperties1.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = cell.Font.Color });
                    //run.Append(runProperties1);
                    //#endregion

                    //#region 문자 간격
                    //Spacing spacing4 = new Spacing() { Val = (int)(cell.Font.CharacterSpacing * 20) };
                    //runProperties1.Append(spacing4);
                    //#endregion

                    #region 텍스트 자동 맞춤
                    if (cell.Paragraph.TableCellFitText)
                    {
                        TableCellFitText cellFitText = new TableCellFitText();
                        runProperties1.Append(cellFitText);
                    }
                    #endregion
                    run.Append(runProperties1);


                    #region 텍스트
                    //워드 텍스트는 개행이 개판이다. 
                    //PPT는 \n 이 자동 처리된다.
                    if (cell.Paragraph.TextDirection != Model.Component.HomeTab.TextDirection.Stacked)
                    {
                        string[] strs = cell.Text.Split('\n');
                        for (int i = 0; i < strs.Length; i++)
                        {
                            Text text1 = new Text();
                            text1.Text = strs[i];
                            text1.Space = SpaceProcessingModeValues.Preserve;
                            run.Append(text1);

                            if (i < strs.Length - 1)
                            {
                                run.AppendChild(new Break());
                            }
                        }
                    }
                    else
                    {
                        Text text1 = new Text();
                        text1.Space = SpaceProcessingModeValues.Preserve;
                        text1.Text = cell.Text;
                        run.Append(text1);
                    }
                    #endregion

                    return run;
                }
            }

            TableCellProperties SetCellProperty()
            {
                var tableCellProperties = new TableCellProperties();

                #region 셀병합
                if(isSingleCell == false)
                {
                    if (cell.TableInfo.Cell.RowSpan > 1)
                    {
                        VerticalMerge verticalMerge = new VerticalMerge();
                        verticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                        tableCellProperties.Append(verticalMerge);
                    }
                    if (cell.TableInfo.Cell.ColSpan > 1)
                    {
                        GridSpan gridSpan = new GridSpan() { Val = cell.TableInfo.Cell.ColSpan };
                        tableCellProperties.Append(gridSpan);
                    }
                }

                #endregion

                #region 여백
                tableCellProperties.Append(SetPropertyMargin(cell.Margin.Left, cell.Margin.Right, cell.Margin.Top,cell.Margin.Bottom));
                #endregion

                #region 테두리

                tableCellProperties.Append(GetTableCellBorders(cell.TableInfo.Styles));
                #endregion

                #region 배경색
                if (cell.TableInfo.Styles.UseShading == false)
                {
                    //??
                }
                else
                {
                    Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = cell.TableInfo.Styles.ShadingColor };
                    tableCellProperties.Append(shading1);
                }
                #endregion

                #region 세로정렬
                TableCellVerticalAlignment vAlignment = new TableCellVerticalAlignment()
                {
                    Val = Common.Instance.GetWordprocessingTableVerticalAlignment(cell.Paragraph.AlignmentVertical)
                };
                tableCellProperties.Append(vAlignment);
                #endregion

                #region 텍스트 가로 세로 여부
                TextDirection textDirection1 = new TextDirection()
                {
                    Val = Common.Instance.GetWordpressingTextDirection(cell.Paragraph.TextDirection)
                };
                tableCellProperties.Append(textDirection1);
                #endregion

                return tableCellProperties;
            }
        }


        TableCellMargin SetPropertyMargin(float left, float right, float top, float bottom)
        {
            TableCellMargin margin = new TableCellMargin();
            margin.LeftMargin = new LeftMargin() { Width = (left *wordTableRatio).ToString() };
            margin.RightMargin = new RightMargin() { Width = (right * wordTableRatio).ToString() };
            margin.TopMargin = new TopMargin() { Width = (top * wordTableRatio).ToString() };
            margin.BottomMargin = new BottomMargin() { Width = (bottom * wordTableRatio).ToString() };

            return margin;
        }

        /// <summary>
        /// 빈 셀을 생성합니다. 워드의 경우 병합 Column은 생성하지 않고, 병합 Row만 생성됩니다. 
        /// </summary>
        TableCell GenerateEmptyTableCell(OfficeModel cell)
        {
            TableCell tableCell1 = new TableCell();
            TableCellProperties tableCellProperties1 = new TableCellProperties();
            VerticalMerge verticalMerge1 = new VerticalMerge();
            if (cell.TableInfo.Cell.ColSpan > 1)
            {
                GridSpan gridSpan = new GridSpan() { Val = cell.TableInfo.Cell.ColSpan };
                tableCellProperties1.Append(gridSpan);
            }
            tableCellProperties1.Append(GetTableCellBorders(cell.TableInfo.Styles));

            tableCellProperties1.Append(verticalMerge1);
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00DF65E6", RsidRunAdditionDefault = "00DF65E6" };

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);
            return tableCell1;
        }

        /// <summary>
        /// 테이블 셀 테두리를 가져옵니다. 
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        TableCellBorders GetTableCellBorders(OfficeTableStyles style)
        {
            TableCellBorders tableCellBorders1 = new TableCellBorders();
            //! 주의! 컨트롤 한개씩 그려지는게 아니라. 일괄로 좌측 그리고, 우측그리고 상단 그리고 하단 그리고 하는것 같다. 
            if ((int)style.Left.Weight > 0)
            {
                int size = style.Left.Draw ? (int)style.Left.Weight : 0;

                //todo border 스타일 구현안됨.
                LeftBorder border = new LeftBorder()
                {
                    Val = BorderValues.Single,
                    Color = style.Left.Color,
                    Size = Convert.ToUInt32(4 * size),
                    Space = 0U
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Right.Weight > 0)
            {
                int size = style.Right.Draw ? (int)style.Right.Weight : 0;
                RightBorder border = new RightBorder()
                {
                    Val = BorderValues.Single,
                    Color = style.Right.Color,
                    Size = Convert.ToUInt32(4 * size),
                    Space = 0U
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Top.Weight > 0)
            {
                int size = style.Top.Draw ? (int)style.Top.Weight : 0;
                TopBorder border = new TopBorder()
                {
                    Val = BorderValues.Single,
                    Color = style.Top.Color,
                    Size = Convert.ToUInt32(4 * size),
                    Space = 0U
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Bottom.Weight > 0)
            {
                int size = style.Bottom.Draw ? (int)style.Bottom.Weight : 0;
                BottomBorder border = new BottomBorder()
                {
                    Val = BorderValues.Single,
                    Color = style.Bottom.Color,
                    Size = Convert.ToUInt32(4 * size),
                    Space = 0U
                };
                tableCellBorders1.Append(border);
            }
            return tableCellBorders1;
        }
    }
}
