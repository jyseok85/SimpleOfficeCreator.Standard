using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    internal class SocAlternateContent
    {

        private SocAlternateContent() { }
        //private static 인스턴스 객체
        private static readonly Lazy<SocAlternateContent> _instance = new Lazy<SocAlternateContent>(() => new SocAlternateContent());
        //public static 의 객체반환 함수
        public static SocAlternateContent Instance { get { return _instance.Value; } }

        public AlternateContent GetAlternateContent(OfficeModel model)
        {
            //var a = new testclass();
            //return a.GenerateAlternateContent(); 

            AlternateContent alternateContent = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing1 = new Drawing();

            #region 1.마진
            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0, DistanceFromBottom = (UInt32Value)0, DistanceFromLeft = (UInt32Value)0, DistanceFromRight = (UInt32Value)0, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
            //Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)45720U, DistanceFromBottom = (UInt32Value)45720U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251664384U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "6FC46515", AnchorId = "15C98A1E" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };
            #endregion

            #region 2.위치
            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = (model.Rect.X * 9525).ToString();
            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = (model.Rect.Y * 9525).ToString();
            verticalPosition1.Append(positionOffset2);
            #endregion

            #region 3.크기
            Wp.Extent extent1 = new Wp.Extent() { Cx = model.Rect.Width * 9525, Cy = model.Rect.Height * 9525 };
            //Wp.Extent extent1 = new Wp.Extent() { Cx = 2360930L, Cy = 474453L };
            #endregion

            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 22860L, BottomEdge = 11430L };
            //Wp.WrapSquare wrapSquare1 = new Wp.WrapSquare() { WrapText = Wp.WrapTextValues.BothSides };

            #region ID 설정
            uint uniqueId = Common.Instance.UniqueId.Last() + 1;
            Common.Instance.UniqueId.Add(uniqueId);
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)uniqueId, Name = model.Name };
            #endregion


            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };



            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);


          
            


            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();



            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };
            nonVisualDrawingShapeProperties1.Append(shapeLocks1);


            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            #region 도형타입
            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();
            presetGeometry1.Append(adjustValueList1);
            shapeProperties1.Append(presetGeometry1);
            #endregion

            #region 배경색
            if (model.ShapeStyle.UseFill == false)
            {
                shapeProperties1.Append(new A.NoFill());
            }
            else
            {
                shapeProperties1.Append(Common.Instance.GenerateSolidFill(model.ShapeStyle.FillColor));
            }

            #endregion

            #region 테두리
            if (model.ShapeStyle.UseOutline && model.ShapeStyle.OutlineWeight > 0)
            {
                var outline = Common.Instance.GetDrawingOutline(model.ShapeStyle.OutlineWeight, model.ShapeStyle.OutlineColor);
                shapeProperties1.Append(outline);
            }
            //A.Outline outline1 = new A.Outline() { Width = 9525 };
            //A.SolidFill solidFill1 = new A.SolidFill();
            //A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            //solidFill1.Append(rgbColorModelHex1);

            //A.SolidFill solidFill2 = new A.SolidFill();
            //A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };
            //solidFill2.Append(rgbColorModelHex2);
            //A.Miter miter1 = new A.Miter() { Limit = 800000 };
            //A.HeadEnd headEnd1 = new A.HeadEnd();
            //A.TailEnd tailEnd1 = new A.TailEnd();

            //outline1.Append(solidFill2);
            //outline1.Append(miter1);
            //outline1.Append(headEnd1);
            //outline1.Append(tailEnd1);

            //shapeProperties1.Append(Common.Instance.GetDrawingTransfrom2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));
            //shapeProperties1.Append(solidFill1);
            //shapeProperties1.Append(outline1);
            #endregion

            #region 내부 편집 컨트롤 
            shapeProperties1.Append(Common.Instance.GetDrawingTransfrom2D(0,0, model.Rect.Width, model.Rect.Height));
            #endregion


            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            #region 텍스트 수정불가 블록(삭제가능)
            //SdtBlock sdtBlock1 = new SdtBlock();

            //#region 도형내부의 컨트롤
            //SdtProperties sdtProperties1 = new SdtProperties();
            //SdtId sdtId1 = new SdtId() { Val = 568603642 };
            //TemporarySdt temporarySdt1 = new TemporarySdt();
            //ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
            //W15.Appearance appearance1 = new W15.Appearance() { Val = W15.SdtAppearance.Hidden };

            //sdtProperties1.Append(sdtId1);
            //sdtProperties1.Append(temporarySdt1);
            //sdtProperties1.Append(showingPlaceholder1);
            //sdtProperties1.Append(appearance1);
            //#endregion

            //SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            //Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00573810", RsidRunAdditionDefault = "00573810" };

            //Run run2 = new Run();
            //Text text1 = new Text(); 
            //text1.Text = model.Text;
            //run2.Append(text1);

            //paragraph1.Append(run2);

            //sdtContentBlock1.Append(paragraph1);

            //sdtBlock1.Append(sdtProperties1);
            //sdtBlock1.Append(sdtContentBlock1);

            //textBoxContent1.Append(sdtBlock1);
            #endregion

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "004B59EC", RsidParagraphAddition = "007669DE", RsidParagraphProperties = "00060CA5", RsidRunAdditionDefault = "007669DE" };


            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            //Bold bold1 = new Bold();
            //Italic italic1 = new Italic();
            //Color color1 = new Color() { Val = "FF0000" };
            //Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            //paragraphMarkRunProperties1.Append(bold1);
            //paragraphMarkRunProperties1.Append(italic1);
            //paragraphMarkRunProperties1.Append(underline1);
            //paragraphMarkRunProperties1.Append(color1);
            RunFonts runFonts103 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            paragraphMarkRunProperties1.Append(runFonts103);
            paragraphProperties1.Append(paragraphMarkRunProperties1);


            Run run1 = new Run();
            //RunProperties runProperties1 = new RunProperties();
            //Bold bold2 = new Bold();
            //Italic italic2 = new Italic();
            //Color color2 = new Color() { Val = "FF0000" };
            //Highlight highlight1 = new Highlight() { Val = HighlightColorValues.DarkBlue };
            //Underline underline2 = new Underline() { Val = UnderlineValues.Single };

            //runProperties1.Append(bold2);
            //runProperties1.Append(italic2);
            //runProperties1.Append(color2);
            //runProperties1.Append(highlight1);
            //runProperties1.Append(underline2);

            Text text1 = new Text();
            text1.Text = model.Text;

            //var runProperties = Common.Instance.GetWordRunProperty(model.Font);
            //run1.Append(runProperties);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            textBoxContent1.Append(paragraph1);
            textBoxInfo21.Append(textBoxContent1);

            //Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties(){ Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 28800, TopInset = 0, RightInset = 28800, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false };
            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties()
            {
                Rotation = 0,
                Vertical = Common.Instance.GetDrawingTextVertical(model.Paragraph.TextDirection),
                Wrap = A.TextWrappingValues.Square, //이걸 적용해야 지정한 사이즈에 딱맞게 생성된다. ppt는 없는듯한데..
                LeftInset = (int)model.Margin.Left * Common.Instance.EMUPPI,
                TopInset = (int)model.Margin.Top * Common.Instance.EMUPPI,
                RightInset = (int)model.Margin.Right * Common.Instance.EMUPPI,
                BottomInset = (int)model.Margin.Bottom * Common.Instance.EMUPPI,
                Anchor = Common.Instance.GetDrawingAnchoring(model.Paragraph.AlignmentVertical),
                AnchorCenter = false
            };

            #region 도형을 텍스트크기에 맞춤.
            //A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();
            //textBodyProperties1.Append(shapeAutoFit1);
            #endregion

            #region 도형을 텍스트크기에 안맞춤.
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();
            textBodyProperties1.Append(noAutoFit1);
            #endregion

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);


            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(new Wp.WrapNone());
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            alternateContentChoice1.Append(drawing1);

            alternateContent.Append(alternateContentChoice1);


            return alternateContent;
        }
    }
}
