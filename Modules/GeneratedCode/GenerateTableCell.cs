using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.Model;
using SimpleOfficeCreator.Stardard.Modules.Model.Component;
using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class GenerateTableCell
    {
        private GenerateTableCell() { }
        //private static 인스턴스 객체
        private static readonly Lazy<GenerateTableCell> _instance = new Lazy<GenerateTableCell>(() => new GenerateTableCell());
        //public static 의 객체반환 함수
        public static GenerateTableCell Instance { get { return _instance.Value; } }

        public int EMUPPI { get; set; } = 0;
        //DR은 기본 일반 양식보다 자간을 적게 사용한다. 그러므로 일정부분 작게 만든다.
        public A.Graphic Graphic(OfficeModel model)
        {
            var  graphic1 = new A.Graphic();
            var  graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };

            graphicData1.Append(GenerateTable(model));

            graphic1.Append(graphicData1);
            return graphic1;

        }

        private A.PresetLineDashValues GetDrawingDashValue(string style)
        {
            A.PresetLineDashValues dashValue = A.PresetLineDashValues.Solid;

            switch (style.ToUpper())
            {
                case "SOLID":
                    dashValue = A.PresetLineDashValues.Solid; break;
                case "DOT":
                    dashValue = A.PresetLineDashValues.Dot; break;
                case "DASH":
                    dashValue = A.PresetLineDashValues.Dash; break;
                case "DASHDOT":
                    dashValue = A.PresetLineDashValues.DashDot; break;
                case "DASHDOTDOT":
                    dashValue = A.PresetLineDashValues.SystemDashDotDot; break;
                default:
                    break;

            }

            return dashValue;
        }


        private A.Table GenerateTable(OfficeModel model)
        {  
            //속성 적용할지 미정(단어 중간에 글자 짤림가능 기능)
            bool lineBreak = true;         

            var table1 = new A.Table();
            table1.Append(GenerateTableProperties());
            table1.Append(GenerateTableGrid());
            if(model.TableInfo is null)
                table1.Append(GenerateTableRow(model.Rect.Height, -1));
            else
            {
                for(int i = 0; i < model.TableInfo.RowHeightList.Count; i++ )
                {
                    table1.Append(GenerateTableRow(model.TableInfo.RowHeightList[i], i));
                }
            }
            return table1;

            A.TableProperties GenerateTableProperties()
            {
                A.TableProperties tableProperties1 = new A.TableProperties();
                A.TableStyleId tableStyleId1 = new A.TableStyleId();
                tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

                tableProperties1.Append(tableStyleId1);
                return tableProperties1;
            }            
            //대표적으로 컬럼 사이즈를 설정한다.
            A.TableGrid GenerateTableGrid()
            {                
                A.TableGrid tableGrid1 = new A.TableGrid();

                if(model.TableInfo is null)
                {
                    A.GridColumn gridColumn1 = new A.GridColumn() { Width = model.Rect.Width * EMUPPI };
                    A.ExtensionList extensionList1 = new A.ExtensionList();
                    extensionList1.Append(new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" });
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
                        A.GridColumn gridColumn1 = new A.GridColumn() { Width = colWidth * EMUPPI };
                        A.ExtensionList extensionList1 = new A.ExtensionList();
                        extensionList1.Append(new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" }); 
                        gridColumn1.Append(extensionList1);
                        tableGrid1.Append(gridColumn1);

                        //Logger.Instance.Write(gridColumn1.Width.ToString());
                    }
                }   
                return tableGrid1;
            }
            A.TableRow GenerateTableRow(int height, int row)
            {
                A.TableRow tableRow1 = new A.TableRow() { Height = height * EMUPPI ,
                    
                    
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
                A.TableCell GeterateEmptyCell(OfficeModel cell)
                {
                    var tableCell1 = new A.TableCell();
                    tableCell1.HorizontalMerge = cell.TableInfo.Cell.HorizontalMerge;
                    tableCell1.VerticalMerge = cell.TableInfo.Cell.VerticalMerge;
                    A.TextBody textBody1 = new A.TextBody();
                    BodyProperties bodyProperties1 = new BodyProperties();
                    ListStyle listStyle1 = new ListStyle();

                    A.Paragraph paragraph1 = new A.Paragraph();
                    A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = TextAlignmentTypeValues.Center, LatinLineBreak = lineBreak };
                    EndParagraphRunProperties endParagraphRunProperties1 = new EndParagraphRunProperties() { FontSize = 800, Dirty = false };

                    paragraph1.Append(paragraphProperties1);
                    paragraph1.Append(endParagraphRunProperties1);

                    textBody1.Append(bodyProperties1);
                    textBody1.Append(listStyle1);
                    textBody1.Append(paragraph1);
                    A.TableCellProperties tableCellProperties1 = new A.TableCellProperties();

                    tableCell1.Append(textBody1);
                    tableCell1.Append(tableCellProperties1);
                    return tableCell1;
                }
                A.TableCell GenerateTableCell(OfficeModel cell)
                {
                    A.TableCell tableCell1 = new A.TableCell();

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
                        A.TextBody textBody1 = new A.TextBody();
                        //단락 : 단락 속성과 텍스트로 구성됨.
                        A.Paragraph paragraph1 = new A.Paragraph();

                        //1. 단락 속성 : 수평 정렬
                        A.TextAlignmentTypeValues textAlignment = cell.Paragraph.AlignmentHorizontal;
                        A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties()
                        {
                            Alignment = textAlignment,
                            LatinLineBreak = lineBreak,
                        };

                        //줄간격 옵션. 그러나 사용안함.
                        //LineSpacing lineSpacing1 = new LineSpacing();
                        //SpacingPoints spacingPoints1 = new SpacingPoints() { Val = (int)(cell.Font.Size * 1.2) * 100 };
                        //lineSpacing1.Append(spacingPoints1);
                        //paragraphProperties1.Append(lineSpacing1);

                        paragraph1.Append(paragraphProperties1);

                        A.Run run1 = new A.Run();

                        //2. 텍스트 속성 : 폰트, 컬러, 내용, Bold 등
                        SetTextProperty();

                        #region [사용안함] end paragraph
                        //이건 뭐냐?? end paragraph? 다음에 삽입되는 텍스트의 속성을 지정한다.
                        //즉.. 이미 만들어진 텍스트 및 속성을 갖고 있는 컨버트에서는 아무 의미가 없다.
                        //A.EndParagraphRunProperties endParagraphRunProperties1 = new()
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
                        //A.SolidFill solidFill2 = new();
                        //A.SchemeColor schemeColor2 = new() { Val = A.SchemeColorValues.Text1 };
                        //solidFill2.Append(schemeColor2);

                        //A.LatinFont latinFont2 = new A.LatinFont() { Typeface = fontFace };
                        //A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = fontFace };

                        //endParagraphRunProperties1.Append(solidFill2);
                        //endParagraphRunProperties1.Append(latinFont2);
                        //endParagraphRunProperties1.Append(eastAsianFont2);
                        #endregion
                        paragraph1.Append(run1);

                        BodyProperties body = new A.BodyProperties()
                        {
                            VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
                            Wrap = A.TextWrappingValues.Square
                            
                        };
                        textBody1.Append(body);
                        //필수
                        //textBody1.Append(new A.BodyProperties());
                        //필수
                        textBody1.Append(new A.ListStyle());
                        textBody1.Append(paragraph1);
                        tableCell1.Append(textBody1);
                        
                        void SetTextProperty()
                    {
                        A.TextUnderlineValues underlineValue = A.TextUnderlineValues.None;
                        if (cell.Font.UnderLine)
                            underlineValue = A.TextUnderlineValues.Single;

                        A.TextStrikeValues strikeValues = A.TextStrikeValues.NoStrike;
                        if (cell.Font.Strike)
                            strikeValues = A.TextStrikeValues.SingleStrike;                        

                        A.RunProperties runProperties1 = new A.RunProperties()
                        {
                            Language = "en-US",
                            AlternativeLanguage = "ko-KR",
                            Dirty = false,
                            FontSize = (int)cell.Font.Size * 100,
                            Bold = cell.Font.Bold,
                            Italic = cell.Font.Italic,
                            Underline = underlineValue,
                            Strike = strikeValues,
                            Spacing = (int)(cell.Font.SpacingValue * 100)
                        };

                        A.SolidFill solidFill1 = Common.Instance.GenerateSolidFill(cell.Font.Color);
                        A.LatinFont latinFont1 = new A.LatinFont() { Typeface = cell.Font.Name };
                        A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = cell.Font.Name };


                        runProperties1.Append(solidFill1);
                        runProperties1.Append(latinFont1);
                        runProperties1.Append(eastAsianFont1);


                        A.Text text1 = new A.Text();
                            //텍스트가 없으면 속성이 하나도 안먹는다... 그래서 빈 공백을 추가한다. 이해안감. 버그아님?
                            if (cell.Text == string.Empty)
                                cell.Text = " ";
                            text1.Text = cell.Text;
                               //text1.Text = cell.Rect.Width.ToString();

                            run1.Append(runProperties1);
                        run1.Append(text1);
                    }
                    }

                    void SetCellProperty()
                    {
                        //텍스트 수직정렬(방향)
                        A.TextVerticalValues textVerticalValue = A.TextVerticalValues.Horizontal;
                        switch (cell.Paragraph.TextDirection)
                        {
                            case Model.Component.HomeTab.TextDirection.Vertical:
                                textVerticalValue = A.TextVerticalValues.EastAsianVetical;
                                break;
                            case Model.Component.HomeTab.TextDirection.RotateAllText90:
                                textVerticalValue = A.TextVerticalValues.Vertical;
                                break;
                            case Model.Component.HomeTab.TextDirection.RotateAllText270:
                                textVerticalValue = A.TextVerticalValues.Vertical270;
                                break;
                            case Model.Component.HomeTab.TextDirection.Stacked:
                                textVerticalValue = A.TextVerticalValues.WordArtVertical;
                                break;
                        }
                        A.TextAnchoringTypeValues textAnchoring = cell.Paragraph.AlignmentVertical;
                        A.TableCellProperties tableCellProperties1 = new A.TableCellProperties()
                        {
                            Vertical = textVerticalValue,
                            Anchor = textAnchoring,
                            LeftMargin = (int)cell.Margin.Left * EMUPPI,
                            RightMargin = (int)cell.Margin.Right * EMUPPI,
                            TopMargin = (int)cell.Margin.Top * EMUPPI,
                            BottomMargin = (int)cell.Margin.Bottom * EMUPPI
                        };

                        
                        //! 주의! 컨트롤 한개씩 그려지는게 아니라. 일괄로 좌측 그리고, 우측그리고 상단 그리고 하단 그리고 하는것 같다. 
                        if ((int)cell.TableInfo.Styles.Left.Weight > 0)
                        {
                            A.LeftBorderLineProperties leftBorderLineProperties1 = new LeftBorderLineProperties()
                            {
                                Width = (int)cell.TableInfo.Styles.Left.Weight * EMUPPI,
                            };
                            AddBorder(leftBorderLineProperties1, 
                                      GetDrawingDashValue(cell.TableInfo.Styles.Left.Style), 
                                      cell.TableInfo.Styles.Left.Draw, 
                                      cell.TableInfo.Styles.Left.Color);
                            tableCellProperties1.Append(leftBorderLineProperties1);
                        }
                        if ((int)cell.TableInfo.Styles.Right.Weight > 0)
                        {
                            A.RightBorderLineProperties rightBorderLineProperties1 = new RightBorderLineProperties()
                            {
                                Width = (int)cell.TableInfo.Styles.Right.Weight * EMUPPI,
                            };
                            AddBorder(rightBorderLineProperties1, GetDrawingDashValue(cell.TableInfo.Styles.Right.Style), cell.TableInfo.Styles.Right.Draw, cell.TableInfo.Styles.Right.Color);
                            tableCellProperties1.Append(rightBorderLineProperties1);
                        }
                        if ((int)cell.TableInfo.Styles.Top.Weight > 0)
                        {
                            A.TopBorderLineProperties topBorderLineProperties1 = new TopBorderLineProperties()
                            {
                                Width = (int)cell.TableInfo.Styles.Top.Weight * EMUPPI,
                            };
                            AddBorder(topBorderLineProperties1, GetDrawingDashValue(cell.TableInfo.Styles.Top.Style), cell.TableInfo.Styles.Top.Draw, cell.TableInfo.Styles.Top.Color);
                            tableCellProperties1.Append(topBorderLineProperties1);
                        }
                        if ((int)cell.TableInfo.Styles.Bottom.Weight > 0)
                        {
                            A.BottomBorderLineProperties bottomBorderLineProperties1 = new BottomBorderLineProperties()
                            {
                                Width = (int)cell.TableInfo.Styles.Bottom.Weight * EMUPPI,
                            };
                            AddBorder(bottomBorderLineProperties1, GetDrawingDashValue(cell.TableInfo.Styles.Bottom.Style), cell.TableInfo.Styles.Bottom.Draw, cell.TableInfo.Styles.Bottom.Color);
                            tableCellProperties1.Append(bottomBorderLineProperties1);
                        }


                            



                        ///대각선 그리기(형식만 만듬. DR에서 대각선이 없기 때문에 실제 사용되지는 않음)
                        DrawDiagonal();
                        #region 배경색
                        if (cell.TableInfo.Styles.UseShading == false)
                        {
                            tableCellProperties1.Append(new A.NoFill());
                        }
                        else
                        {
                            //배경색 추가
                            tableCellProperties1.Append(Common.Instance.GenerateSolidFill(cell.TableInfo.Styles.ShadingColor));
                        }
                        #endregion

                        tableCell1.Append(tableCellProperties1);

                        void AddBorder(A.LinePropertiesType line, A.PresetLineDashValues style, bool isDraw, string color)
                        {
                            line.CapType = A.LineCapValues.Flat;
                            line.CompoundLineType = A.CompoundLineValues.Single;
                            line.Alignment = A.PenAlignmentValues.Center;
                            if (isDraw)
                            {
                                line.Append(Common.Instance.GenerateSolidFill(color));
                            }
                            else
                            {
                                A.NoFill noFill1 = new NoFill();
                                line.Append(noFill1);
                            }


                            var presetDash2 = new PresetDash() { Val = style };
                            var round2 = new Round();
                            var headEnd2 = new HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
                            var tailEnd2 = new TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };


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
                            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
                            A.NoFill noFill2 = new NoFill();
                            A.PresetDash presetDash5 = new PresetDash() { Val = A.PresetLineDashValues.Solid };

                            topLeftToBottomRightBorderLineProperties1.Append(noFill2);
                            topLeftToBottomRightBorderLineProperties1.Append(presetDash5);

                            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
                            A.NoFill noFill3 = new NoFill();
                            A.PresetDash presetDash6 = new PresetDash() { Val = A.PresetLineDashValues.Solid };

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

        public Transform Transform(int x, int y, int width, int height)
        {
            Transform transform1 = new Transform();
            A.Offset offset1 = new Offset() { X = x * EMUPPI, Y = y * EMUPPI };
            A.Extents extents1 = new Extents() { Cx = width * EMUPPI, Cy = height * EMUPPI };

            transform1.Append(offset1);
            transform1.Append(extents1);
            return transform1;
        }
    }
}
