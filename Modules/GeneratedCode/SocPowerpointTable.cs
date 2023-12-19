using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using SimpleOfficeCreator.Standard.Modules.Model;
using System;
using System.Collections.Generic;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using PPT = DocumentFormat.OpenXml.Presentation;

namespace SimpleOfficeCreator.Standard.Modules.GeneratedCode
{
    public class SocPowerpointTable
    {
        private SocPowerpointTable() { }
        //private static 인스턴스 객체
        private static readonly Lazy<SocPowerpointTable> _instance = new Lazy<SocPowerpointTable>(() => new SocPowerpointTable());
        //public static 의 객체반환 함수
        public static SocPowerpointTable Instance { get { return _instance.Value; } }

        //public PPT.GraphicFrame GenerateGraphicFrame(OfficeModel model)
        //{
        //    var graphicFrame = new PPT.GraphicFrame();
        //    graphicFrame.Append(StaticCode.GenerateNonVisualGraphicFrameProperties("표"));

        //    PPT.Transform transform1 = SocPowerpointTable.Instance.Transform(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height);
        //    graphicFrame.Append(transform1);

        //    Graphic graphic1 = SocPowerpointTable.Instance.Graphic(model);
        //    //A.Graphic graphic1 = Graphic(model);
        //    graphicFrame.Append(graphic1);

        //    return graphicFrame;
        //}


        //public Graphic Graphic(OfficeModel model)
        //{
        //    var graphic1 = new Graphic();
        //    var graphicData1 = new GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };

        //    graphicData1.Append(GenerateTable(model));

        //    graphic1.Append(graphicData1);
        //    return graphic1;

        //}
        public PPT.GraphicFrame Generate(OfficeModel model)
        {
            var graphicFrame = new PPT.GraphicFrame();
            graphicFrame.Append(GetNonVisualGraphicFrameProperties("표"));

            var transform1 = Transform(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height);
            graphicFrame.Append(transform1);

            var graphic1 = new Graphic();
            var graphicData1 = new GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
            graphicData1.Append(GenerateTable(model));
            graphic1.Append(graphicData1);
            graphicFrame.Append(graphic1);

            return graphicFrame;

            PPT.NonVisualGraphicFrameProperties GetNonVisualGraphicFrameProperties(string tableName)
            {
                var nonVisualGraphicFrameProperties1 = new PPT.NonVisualGraphicFrameProperties();
                var nonVisualDrawingProperties1 = new PPT.NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = tableName };

                var nonVisualGraphicFrameDrawingProperties1 = new PPT.NonVisualGraphicFrameDrawingProperties();
                GraphicFrameLocks graphicFrameLocks1 = new GraphicFrameLocks() { NoGrouping = true };

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                var applicationNonVisualDrawingProperties1 = new PPT.ApplicationNonVisualDrawingProperties();

                var applicationNonVisualDrawingPropertiesExtensionList1 = new PPT.ApplicationNonVisualDrawingPropertiesExtensionList();

                var applicationNonVisualDrawingPropertiesExtension1 = new PPT.ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

                P14.ModificationId modificationId1 = new P14.ModificationId() { Val = (UInt32Value)959388604U };
                modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

                applicationNonVisualDrawingPropertiesExtension1.Append(modificationId1);

                applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

                applicationNonVisualDrawingProperties1.Append(applicationNonVisualDrawingPropertiesExtensionList1);

                nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
                nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
                nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties1);
                return nonVisualGraphicFrameProperties1;
            }

        }

        public PPT.Transform Transform(int x, int y, int width, int height)
        {
            PPT.Transform transform1 = new PPT.Transform();
            Offset offset1 = new Offset() { X = x * Common.Instance.EMUPPI, Y = y * Common.Instance.EMUPPI };
            Extents extents1 = new Extents() { Cx = width * Common.Instance.EMUPPI, Cy = height * Common.Instance.EMUPPI };

            transform1.Append(offset1);
            transform1.Append(extents1);
            return transform1;
        }


        private Table GenerateTable(OfficeModel model)
        {
            var table1 = new Table();
            table1.Append(GenerateTableProperties());
            table1.Append(GenerateTableGrid(model));
            if (model.TableInfo is null)
                table1.Append(GenerateTableRow(model.Rect.Height, -1, model));
            else
            {
                for (int i = 0; i < model.TableInfo.RowHeightList.Count; i++)
                {
                    table1.Append(GenerateTableRow(model.TableInfo.RowHeightList[i], i, model));
                }
            }
            return table1;


        }

        TableProperties GenerateTableProperties()
        {
            TableProperties tableProperties1 = new TableProperties();
            TableStyleId tableStyleId1 = new TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
            tableProperties1.Append(tableStyleId1);
            return tableProperties1;
        }
        //대표적으로 컬럼 사이즈를 설정한다.
        TableGrid GenerateTableGrid(OfficeModel model)
        {
            TableGrid tableGrid1 = new TableGrid();

            if (model.TableInfo is null)
            {
                GridColumn gridColumn1 = new GridColumn() { Width = model.Rect.Width * Common.Instance.EMUPPI };
                ExtensionList extensionList1 = new ExtensionList();
                extensionList1.Append(new Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" });
                gridColumn1.Append(extensionList1);
                tableGrid1.Append(gridColumn1);
            }
            else
            {
                //X좌표수만큼 컬럼을 만든다. 
                foreach (int colWidth in model.TableInfo.ColumnWidthList)
                {
                    //병합이 없을 경우는 상관없는데, 병합이 있을 경우 컬럼의 최소사이즈가 100000 이상 되어야 한다. 
                    //100000미만 일경우 PPT 실행시 자동으로 증가되어서 테이블사이즈가 틀어진다. 
                    //..... 증상찾는데 하루 꼬박 사용.
                    GridColumn gridColumn1 = new GridColumn() { Width = colWidth * Common.Instance.EMUPPI };
                    ExtensionList extensionList1 = new ExtensionList();
                    extensionList1.Append(new Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" });
                    gridColumn1.Append(extensionList1);
                    tableGrid1.Append(gridColumn1);

                    //Logger.Instance.Write(gridColumn1.Width.ToString());
                }
            }
            return tableGrid1;
        }
        TableRow GenerateTableRow(int height, int row, OfficeModel model)
        {
            TableRow tableRow1 = new TableRow()
            {
                Height = height * Common.Instance.EMUPPI,


            };

            if (row == -1)
            {
                var tableCell = GenerateTableCell(model);
                tableRow1.Append(tableCell);
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
                            tableRow1.Append(GeterateEmptyCell(item));
                        }
                        else
                        {
                            var tableCell = GenerateTableCell(item);
                            tableRow1.Append(tableCell);
                        }
                    }
                }
            }


            return tableRow1;

        }
        TableCell GeterateEmptyCell(OfficeModel cell)
        {
            var tableCell1 = new TableCell();
            tableCell1.HorizontalMerge = cell.TableInfo.Cell.HorizontalMerge;
            tableCell1.VerticalMerge = cell.TableInfo.Cell.VerticalMerge;
            TextBody textBody1 = new TextBody();
            BodyProperties bodyProperties1 = new BodyProperties();
            ListStyle listStyle1 = new ListStyle();

            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties() { Alignment = TextAlignmentTypeValues.Center };
            EndParagraphRunProperties endParagraphRunProperties1 = new EndParagraphRunProperties() { FontSize = 800, Dirty = false };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);
            TableCellProperties tableCellProperties1 = new TableCellProperties();

            tableCell1.Append(textBody1);
            tableCell1.Append(tableCellProperties1);
            return tableCell1;
        }
        TableCell GenerateTableCell(OfficeModel cell)
        {
            TableCell tableCell1 = new TableCell();

            if (cell.TableInfo.Cell.RowSpan > 1)
            {
                tableCell1.RowSpan = cell.TableInfo.Cell.RowSpan;
            }
            if (cell.TableInfo.Cell.ColSpan > 1)
            {
                tableCell1.GridSpan = cell.TableInfo.Cell.ColSpan;

            }

            //1. 텍스트 박스 설정
            SetTextBody();
            //2. 테이블 셀 모양설정 : 텍스트 방향, 수직 정렬, 배경색, 테두리 모양 등
            SetCellProperty();

            return tableCell1;

            void SetTextBody()
            {
                //단락 : 단락 속성과 텍스트로 구성됨.
                Paragraph paragraph1 = new Paragraph();

                //1. 단락 속성 : 수평 정렬
                ParagraphProperties paragraphProperties1 = new ParagraphProperties()
                {
                    //단락 속성 : 수평 정렬
                    Alignment = Common.Instance.GetDrawingAlignment(cell.Paragraph.AlignmentHorizontal),
                    //속성 적용할지 미정(단어 중간에 글자 짤림가능 기능)
                    LatinLineBreak = true,
                };

                #region 줄간격
                if (cell.Paragraph.LineSpacing > 0)
                {
                    LineSpacing lineSpacing1 = new LineSpacing();
                    SpacingPoints spacingPoints1 = new SpacingPoints() { Val = (int)(cell.Paragraph.LineSpacing * 100) };
                    lineSpacing1.Append(spacingPoints1);
                    paragraphProperties1.Append(lineSpacing1);
                }
                #endregion

                paragraph1.Append(paragraphProperties1);

                //2. 텍스트 속성 : 폰트, 컬러, 내용, Bold 등
                #region Run
                var run1 = new Run();
                var runProperties = Common.Instance.GetDrawingRunProperty(cell.Font);
                run1.Append(runProperties);

                run1.Append(Common.Instance.GetDrawingRunText(cell));
                #endregion


                #region [사용안함] end paragraph
                //이건 뭐냐?? end paragraph? 다음에 삽입되는 텍스트의 속성을 지정한다.
                //즉.. 이미 만들어진 텍스트 및 속성을 갖고 있는 컨버트에서는 아무 의미가 없다.
                //EndParagraphRunProperties endParagraphRunProperties1 = new()
                //{
                //    Language = "ko-KR",
                //    AlternativeLanguage = "en-US",
                //    Dirty = false,
                //    FontSize = fontSize,
                //    Bold = bold,
                //    Italic = italic,
                //    Underline = underlineValue,
                //    Strike = strikeValues

                //};
                //SolidFill solidFill2 = new();
                //SchemeColor schemeColor2 = new() { Val = SchemeColorValues.Text1 };
                //solidFill2.Append(schemeColor2);

                //LatinFont latinFont2 = new LatinFont() { Typeface = fontFace };
                //EastAsianFont eastAsianFont2 = new EastAsianFont() { Typeface = fontFace };

                //endParagraphRunProperties1.Append(solidFill2);
                //endParagraphRunProperties1.Append(latinFont2);
                //endParagraphRunProperties1.Append(eastAsianFont2);
                #endregion
                paragraph1.Append(run1);

                TextBody textBody1 = new TextBody();
                BodyProperties body = new BodyProperties()
                {
                    VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
                    Wrap = TextWrappingValues.Square
                };
                textBody1.Append(body);
                //필수
                //textBody1.Append(new BodyProperties());
                //필수
                textBody1.Append(new ListStyle());
                textBody1.Append(paragraph1);
                tableCell1.Append(textBody1);


            }

            void SetCellProperty()
            {
                TableCellProperties tableCellProperties1 = new TableCellProperties()
                {
                    //텍스트 방향
                    Vertical = Common.Instance.GetDrawingTextVertical(cell.Paragraph.TextDirection),
                    //텍스트 수직정렬
                    Anchor = Common.Instance.GetDrawingAnchoring(cell.Paragraph.AlignmentVertical),

                    #region 여백설정
                    LeftMargin = (int)(cell.Margin.Left * Common.Instance.EMUPPI),
                    RightMargin = (int)(cell.Margin.Right * Common.Instance.EMUPPI),
                    TopMargin = (int)(cell.Margin.Top * Common.Instance.EMUPPI),
                    BottomMargin = (int)(cell.Margin.Bottom * Common.Instance.EMUPPI)
                    #endregion
                };



                //! 주의! 컨트롤 한개씩 그려지는게 아니라. 일괄로 좌측 그리고, 우측그리고 상단 그리고 하단 그리고 하는것 같다. 
                if ((int)cell.TableInfo.Styles.Left.Weight > 0)
                {
                    LeftBorderLineProperties leftBorderLineProperties1 = new LeftBorderLineProperties()
                    {
                        Width = (int)cell.TableInfo.Styles.Left.Weight * Common.Instance.EMUPPI,
                    };
                    AddBorder(leftBorderLineProperties1,
                              Common.Instance.GetDrawingDashValue(cell.TableInfo.Styles.Left.Dashes),
                              cell.TableInfo.Styles.Left.Draw,
                              cell.TableInfo.Styles.Left.Color);
                    tableCellProperties1.Append(leftBorderLineProperties1);
                }
                if ((int)cell.TableInfo.Styles.Right.Weight > 0)
                {
                    RightBorderLineProperties rightBorderLineProperties1 = new RightBorderLineProperties()
                    {
                        Width = (int)cell.TableInfo.Styles.Right.Weight * Common.Instance.EMUPPI,
                    };
                    AddBorder(rightBorderLineProperties1,
                        Common.Instance.GetDrawingDashValue(cell.TableInfo.Styles.Right.Dashes), cell.TableInfo.Styles.Right.Draw, cell.TableInfo.Styles.Right.Color);
                    tableCellProperties1.Append(rightBorderLineProperties1);
                }
                if ((int)cell.TableInfo.Styles.Top.Weight > 0)
                {
                    TopBorderLineProperties topBorderLineProperties1 = new TopBorderLineProperties()
                    {
                        Width = (int)cell.TableInfo.Styles.Top.Weight * Common.Instance.EMUPPI,
                    };
                    AddBorder(topBorderLineProperties1, Common.Instance.GetDrawingDashValue(cell.TableInfo.Styles.Top.Dashes), cell.TableInfo.Styles.Top.Draw, cell.TableInfo.Styles.Top.Color);
                    tableCellProperties1.Append(topBorderLineProperties1);
                }
                if ((int)cell.TableInfo.Styles.Bottom.Weight > 0)
                {
                    BottomBorderLineProperties bottomBorderLineProperties1 = new BottomBorderLineProperties()
                    {
                        Width = (int)cell.TableInfo.Styles.Bottom.Weight * Common.Instance.EMUPPI,
                    };
                    AddBorder(bottomBorderLineProperties1, Common.Instance.GetDrawingDashValue(cell.TableInfo.Styles.Bottom.Dashes), cell.TableInfo.Styles.Bottom.Draw, cell.TableInfo.Styles.Bottom.Color);
                    tableCellProperties1.Append(bottomBorderLineProperties1);
                }






                ///대각선 그리기(형식만 만듬. DR에서 대각선이 없기 때문에 실제 사용되지는 않음)
                DrawDiagonal();
                #region 배경색
                if (cell.TableInfo.Styles.UseShading == false)
                {
                    tableCellProperties1.Append(new NoFill());
                }
                else
                {
                    //배경색 추가
                    tableCellProperties1.Append(Common.Instance.GenerateSolidFill(cell.TableInfo.Styles.ShadingColor));
                }
                #endregion

                tableCell1.Append(tableCellProperties1);

                void AddBorder(LinePropertiesType line, PresetLineDashValues style, bool isDraw, string color)
                {
                    line.CapType = LineCapValues.Flat;
                    line.CompoundLineType = CompoundLineValues.Single;
                    line.Alignment = PenAlignmentValues.Center;
                    if (isDraw)
                    {
                        line.Append(Common.Instance.GenerateSolidFill(color));
                    }
                    else
                    {
                        NoFill noFill1 = new NoFill();
                        line.Append(noFill1);
                    }


                    var presetDash2 = new PresetDash() { Val = style };
                    var round2 = new Round();
                    var headEnd2 = new HeadEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium };
                    var tailEnd2 = new TailEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium };


                    line.Append(presetDash2);
                    line.Append(round2);
                    line.Append(headEnd2);
                    line.Append(tailEnd2);
                }
                //대각선 그리기
                void DrawDiagonal()
                {
                    #region 셀 대각선 처리
                    //대각선인듯.
                    TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = CompoundLineValues.Single };
                    NoFill noFill2 = new NoFill();
                    PresetDash presetDash5 = new PresetDash() { Val = PresetLineDashValues.Solid };

                    topLeftToBottomRightBorderLineProperties1.Append(noFill2);
                    topLeftToBottomRightBorderLineProperties1.Append(presetDash5);

                    BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = CompoundLineValues.Single };
                    NoFill noFill3 = new NoFill();
                    PresetDash presetDash6 = new PresetDash() { Val = PresetLineDashValues.Solid };

                    bottomLeftToTopRightBorderLineProperties1.Append(noFill3);
                    bottomLeftToTopRightBorderLineProperties1.Append(presetDash6);


                    tableCellProperties1.Append(topLeftToBottomRightBorderLineProperties1);
                    tableCellProperties1.Append(bottomLeftToTopRightBorderLineProperties1);
                    #endregion
                }
            }
        }
    }
}
