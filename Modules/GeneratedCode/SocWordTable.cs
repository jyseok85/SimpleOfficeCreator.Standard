using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Standard.Modules.Model;
using SimpleOfficeCreator.Standard.Modules.Model.Component.TableDesignTab;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace SimpleOfficeCreator.Standard.Modules.GeneratedCode
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
            bool isSingleTable = model.TableInfo is null;
            Table table = new Table();

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
                    int height = model.TableInfo.RowHeightList[i];
                    TableRow tRow = GenerateTableRow(model, height, i);
                    table.Append(tRow);
                }
            }
            //tableHeight = table.Height;

            return table;
        }


        private TableProperties GenerateTableProperties(bool isTableLabel, PointF loc)
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

        private TableGrid GenerateTableGrid(List<int> columnWidths)
        {
            TableGrid tableGrid = new TableGrid();

            foreach (int colWidth in columnWidths)
            {
                int width = (colWidth) * wordTableRatio;
                GridColumn gridColumn1 = new GridColumn() { Width = width.ToString() };
                tableGrid.Append(gridColumn1);
            }

            return tableGrid;

        }

        private TableRow GenerateTableRow(OfficeModel model, int height, int row)
        {
            TableRow tableRow = new TableRow();
            TableRowProperties tableRowProperties1 = new TableRowProperties();
            tableRowProperties1.Append(new TableRowHeight() { Val = (uint)(height * wordTableRatio), HeightType = HeightRuleValues.Exact });
            tableRow.Append(tableRowProperties1);

            if (row == -1)
            {
                TableCell tableCell = GenerateTableCell(model, true);
                tableRow.Append(tableCell);
            }
            else
            {
                if (model.TableInfo.Children != null)
                {
                    List<OfficeModel> items = model.TableInfo.Children.FindAll(x => x.TableInfo.Cell.Row == row);
                    foreach (OfficeModel item in items)
                    {
                        if (item.TableInfo.Cell.Empty == true)
                        {
                            if (item.TableInfo.Cell.MergedRow == true)
                            {
                                TableCell tableCell = GenerateEmptyTableCell(item);
                                tableRow.Append(tableCell);
                            }
                        }
                        else
                        {
                            TableCell tableCell = GenerateTableCell(item);
                            tableRow.Append(tableCell);
                        }
                    }
                }
            }
            return tableRow;
        }


        private TableCell GenerateTableCell(OfficeModel cell, bool isSingleCell = false)
        {
            TableCell tableCell1 = new TableCell();
            //2. 테이블 셀 모양설정 : 텍스트 방향, 수직 정렬, 배경색, 테두리 모양 등
            tableCell1.Append(SetCellProperty());
            
            //1. 텍스트 박스 설정
            if(cell.TableInfo.Cell.IsImageCell)
            {
                tableCell1.Append(SetPicture());
            }
            else
            {
                tableCell1.Append(SetTextBody());
            }


            return tableCell1;

            Paragraph SetTextBody()
            {
                //단락 : 단락 속성과 텍스트로 구성됨.
                Paragraph paragraph1 = new Paragraph();
                ParagraphProperties paragraphProperties = new ParagraphProperties();
                #region 공백없음 속성
                paragraphProperties.Append(new ParagraphStyleId() { Val = "a3" });
                #endregion

                #region 가로정렬
                paragraphProperties.Append(Common.Instance.GetWordprocessingJustification(cell));
                #endregion

                #region 줄간격
                paragraphProperties.Append(Common.Instance.GetSpacingBetweenLines(cell));
                #endregion
                //순서 중요함. 뒤에 run 프로퍼티보다 늦게 추가되면 안됨.
                paragraph1.Append(paragraphProperties);


                Run run = new Run();
                RunProperties runProperties1 = Common.Instance.GetWordRunProperty(cell.Font);

                #region 텍스트 자동 맞춤(테이블 셀 전용)
                if (cell.Paragraph.TableCellFitText)
                {
                    TableCellFitText cellFitText = new TableCellFitText();
                    runProperties1.Append(cellFitText);
                }
                #endregion

                run.Append(runProperties1);

                #region TEXT
                //워드 텍스트는 개행이 개판이다. 
                //PPT는 \n 이 자동 처리된다.
                Common.Instance.SetWordRunText(run, cell);
                #endregion

                paragraph1.Append(run);

                return paragraph1;


            }


            Paragraph SetPicture()
            {
                Paragraph paragraph1 = new Paragraph();
                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();
                runProperties1.Append(noProof1);

                var darawing = GenerateDrawing(cell);
                run1.Append(runProperties1);
                run1.Append(darawing);
                paragraph1.Append(run1);
                return paragraph1;
            }
            TableCellProperties SetCellProperty()
            {
                TableCellProperties tableCellProperties = new TableCellProperties();

                #region 셀병합
                if (isSingleCell == false)
                {
                    if (cell.TableInfo.Cell.RowSpan > 1)
                    {
                        VerticalMerge verticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
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
                tableCellProperties.Append(SetPropertyMargin(cell.Margin.Left, cell.Margin.Right, cell.Margin.Top, cell.Margin.Bottom));
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

        internal Drawing GenerateDrawing(OfficeModel model)
        {
            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { 
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U, 
                DistanceFromLeft = (UInt32Value)0U, 
                DistanceFromRight = (UInt32Value)0U 
            };
            //Wp.Extent extent1 = new Wp.Extent() { Cx = 5534797L, Cy = 3429479L };

            var pictureModel = Common.Instance.Pictures.Find(x => x.UID == model.UID);
            

            var extent1 = AnchorProperty.Instance.GetExtent(pictureModel.Rect.Width, pictureModel.Rect.Height);
            var effectExtent1 = AnchorProperty.Instance.GetEffectExtent();

            //중복
            uint uniqueId = Common.Instance.UniqueId.Last() + 1;
            Common.Instance.UniqueId.Add(uniqueId);
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)uniqueId, Name = "그림 1" };

            //중복
            #region 그래픽객체
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();
            var graphicFrameLocks1 = new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);


            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = uniqueId, Name = "capture.PNG" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            picture1.Append(nonVisualPictureProperties1);


            Pic.BlipFill blipFill1 = new Pic.BlipFill() { RotateWithShape = true };
            A.Blip blip1 = new A.Blip() { Embed = model.UID, CompressionState = A.BlipCompressionValues.Print };
            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();
            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension1.Append(useLocalDpi1);
            blipExtensionList1.Append(blipExtension1);
            blip1.Append(blipExtensionList1);
            A.Stretch stretch1 = new A.Stretch();

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);
            picture1.Append(blipFill1);


            Pic.ShapeProperties shapeProperties = new Pic.ShapeProperties();
            #region 내부 편집 컨트롤 
            shapeProperties.Append(Common.Instance.GetDrawingTransfrom2D(0, 0, pictureModel.Rect.Width, pictureModel.Rect.Height));
            #endregion

            #region [뭔지모름] 이 요소는 사용자 정의 기하학적 모양 대신 사전 설정된 기하학적 모양을 사용해야 하는 경우를 지정합니다.
            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();
            presetGeometry4.Append(adjustValueList4);
            shapeProperties.Append(presetGeometry4);
            #endregion


            picture1.Append(shapeProperties);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);
            #endregion
            drawing1.Append(inline1);
            return drawing1;

        }

        private TableCellMargin SetPropertyMargin(float left, float right, float top, float bottom)
        {
            TableCellMargin margin = new TableCellMargin
            {
                LeftMargin = new LeftMargin() { Width = (left * wordTableRatio).ToString() },
                RightMargin = new RightMargin() { Width = (right * wordTableRatio).ToString() },
                TopMargin = new TopMargin() { Width = (top * wordTableRatio).ToString() },
                BottomMargin = new BottomMargin() { Width = (bottom * wordTableRatio).ToString() }
            };

            return margin;
        }

        /// <summary>
        /// 빈 셀을 생성합니다. 워드의 경우 병합 Column은 생성하지 않고, 병합 Row만 생성됩니다. 
        /// </summary>
        private TableCell GenerateEmptyTableCell(OfficeModel cell)
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
        private TableCellBorders GetTableCellBorders(OfficeTableStyles style)
        {
            TableCellBorders tableCellBorders1 = new TableCellBorders();
            //! 주의! 컨트롤 한개씩 그려지는게 아니라. 일괄로 좌측 그리고, 우측그리고 상단 그리고 하단 그리고 하는것 같다. 
            if ((int)style.Left.Weight > 0 && style.Left.Color != "transparent")
            {
                int size = style.Left.Draw ? (int)style.Left.Weight : 0;               

                LeftBorder border = new LeftBorder()
                {
                    Val = Common.Instance.GetwordBorderStyle(style.Left.Style),
                    Color = style.Left.Color,
                    Size = Convert.ToUInt32(8 * size),
                    Space = 0
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Right.Weight > 0 && style.Right.Color != "transparent")
            {
                int size = style.Right.Draw ? (int)style.Right.Weight : 0;
                RightBorder border = new RightBorder()
                {
                    Val = Common.Instance.GetwordBorderStyle(style.Right.Style),
                    Color = style.Right.Color,
                    Size = Convert.ToUInt32(8 * size),
                    Space = 0
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Top.Weight > 0 && style.Top.Color != "transparent")
            {
                int size = style.Top.Draw ? (int)style.Top.Weight : 0;
                TopBorder border = new TopBorder()
                {
                    Val = Common.Instance.GetwordBorderStyle(style.Top.Style),
                    Color = style.Top.Color,
                    Size = Convert.ToUInt32(8 * size),
                    Space = 0
                };
                tableCellBorders1.Append(border);
            }
            if ((int)style.Bottom.Weight > 0 && style.Bottom.Color != "transparent")
            {
                int size = style.Bottom.Draw ? (int)style.Bottom.Weight : 0;
                BottomBorder border = new BottomBorder()
                {
                    Val = Common.Instance.GetwordBorderStyle(style.Bottom.Style),
                    Color = style.Bottom.Color,
                    Size = Convert.ToUInt32(8 * size),
                    Space = 0
                };
                tableCellBorders1.Append(border);
            }
            return tableCellBorders1;
        }
    }
}
