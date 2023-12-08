﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;


namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    internal class SocParagraph
    {
        private SocParagraph() { }
        //private static 인스턴스 객체
        private static readonly Lazy<SocParagraph> _instance = new Lazy<SocParagraph>(() => new SocParagraph());
        //public static 의 객체반환 함수
        public static SocParagraph Instance { get { return _instance.Value; } }

        public Paragraph GenerateText(List<OfficeModel> models)
        {
            Paragraph para = new Paragraph();
            foreach (OfficeModel model in models)
            {
                if(model.Type == Model.Type.TextBox)
                    para.Append(GenerateRun(model));
                else if(model.Type == Model.Type.Picture)
                    para.Append(GenerateImage(model));
            }
            return para;
        }

        public Run GenerateImage(OfficeModel model)
        {
            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties();
            //이 요소는 문서의 철자와 문법을 검사할 때 이 실행의 내용이 오류를 보고하지 않도록 지정합니다. 
            NoProof noProof1 = new NoProof();
            runProperties1.Append(noProof1);

            Drawing d = GenerateDrawing(model);
            run1.Append(runProperties1);
            run1.Append(d);
            return run1;
        }

        private Drawing GenerateDrawing(OfficeModel model)// string imageId, int x, int y, int width, int height, bool outLine, float thickness, bool isBackward)
        {
            var drawing1 = new Drawing();

            Anchor anchor1 = new Anchor()
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            SetAnchorProperty(anchor1, model);

            #region 그래픽객체
            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = 3U, Name = "capture.PNG" };
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
            shapeProperties.Append(Common.Instance.GetDrawingTransfrom2D(0, 0, model.Rect.Width, model.Rect.Height));
            #endregion

            #region [뭔지모름] 이 요소는 사용자 정의 기하학적 모양 대신 사전 설정된 기하학적 모양을 사용해야 하는 경우를 지정합니다.
            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();
            presetGeometry4.Append(adjustValueList4);
            shapeProperties.Append(presetGeometry4);
            #endregion

            

            #region 테두리
            if (model.PictureStyle.NoOutline == false && model.PictureStyle.Weight > 0)
            {
                var outline = Common.Instance.GetDrawingOutline(model.PictureStyle.Weight, model.PictureStyle.Color);
                shapeProperties.Append(outline);
            }
            #endregion

            picture1.Append(shapeProperties);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);
            anchor1.Append(graphic1);
            #endregion

            #region [사용안함]상대적크기 그래픽객체를 anchor에 할당후에 설정해야 합니다.(먼저 등록시 오류발생)
            bool useRelativeSize = false;
            if (useRelativeSize)
            {
                Wp14.RelativeWidth relativeWidth4 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
                Wp14.PercentageWidth percentageWidth4 = new Wp14.PercentageWidth();
                percentageWidth4.Text = "0";
                relativeWidth4.Append(percentageWidth4);

                Wp14.RelativeHeight relativeHeight4 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
                Wp14.PercentageHeight percentageHeight4 = new Wp14.PercentageHeight();
                percentageHeight4.Text = "0";
                relativeHeight4.Append(percentageHeight4);

                anchor1.Append(relativeWidth4);
                anchor1.Append(relativeHeight4);
            }
            #endregion

            drawing1.Append(anchor1);
            return drawing1;

        }


        private Run GenerateRun(OfficeModel model)
        {
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            //이 요소는 문서의 철자와 문법을 검사할 때 이 실행의 내용이 오류를 보고하지 않도록 지정합니다. 
            NoProof noProof1 = new NoProof();
            runProperties1.Append(noProof1);

            //var alternateContent1 = SocAlternateContent.Instance.GetAlternateContent(model);
            var alternateContent1 = GetAlternateContent(model);
            run1.Append(runProperties1);
            run1.Append(alternateContent1);
            return run1;
        }

        private void SetAnchorProperty(Anchor anchor, OfficeModel model)
        {
            #region 1.SimplePos
            anchor.Append(AnchorProperty.Instance.GetSimplePosition());
            #endregion

            #region 2.position H
            //기준점설정 현재 페이지의 마진을 설정했기 때문에 Margin을 설정 만약 절대값으로 하고싶으면 Page를 선택한다.
            HorizontalPosition horizontalPosition = new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.Column };
            horizontalPosition.Append(AnchorProperty.Instance.GetPositionOffset(model.Rect.X));
            anchor.Append(horizontalPosition);
            #endregion

            #region 3.position V
            VerticalPosition verticalPosition = new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Paragraph };
            verticalPosition.Append(AnchorProperty.Instance.GetPositionOffset(model.Rect.Y));
            anchor.Append(verticalPosition);
            #endregion

            #region 4.크기 
            anchor.Append(AnchorProperty.Instance.GetExtent(model.Rect.Width, model.Rect.Height));
            #endregion

            #region 5.도형효과
            anchor.Append(AnchorProperty.Instance.GetEffectExtent());
            #endregion

            #region 6.텍스트 줄 바꿈 없음
            anchor.Append(new WrapNone());
            #endregion

            #region 7.문서 속성
            uint uniqueId = Common.Instance.UniqueId.Last() + 1;
            Common.Instance.UniqueId.Add(uniqueId);
            DocProperties docProperties = new DocProperties() { Id = (UInt32Value)uniqueId, Name = model.Name };
            anchor.Append(docProperties);
            #endregion

            #region 8.[사용안함][뭔지모름] 이 요소는 상위 DrawingML 객체에 대한 일반적인 비시각적 DrawingML 객체 속성을 지정합니다
            bool useNonVisualGrapic = true;
            if (useNonVisualGrapic)
            {
                NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties4 = new NonVisualGraphicFrameDrawingProperties();
                A.GraphicFrameLocks graphicFrameLocks4 = new A.GraphicFrameLocks();
                graphicFrameLocks4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                nonVisualGraphicFrameDrawingProperties4.Append(graphicFrameLocks4);
                anchor.Append(nonVisualGraphicFrameDrawingProperties4);
            }
            #endregion

        

        }

        private AlternateContent GetAlternateContent(OfficeModel model)
        {
            var alternateContent = new AlternateContent();
            AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "wps" };
            Drawing drawing4 = new Drawing();

            Anchor anchor = new Anchor() { 
                DistanceFromTop = (UInt32Value)0, 
                DistanceFromBottom = (UInt32Value)0, 
                DistanceFromLeft = (UInt32Value)0, 
                DistanceFromRight = (UInt32Value)0, 
                SimplePos = false, 
                RelativeHeight = (UInt32Value)251659264U, 
                BehindDoc = false, 
                Locked = false, 
                LayoutInCell = true, 
                AllowOverlap = true 
            };


            SetAnchorProperty(anchor, model);

            anchor.Append(SetText(model));

            #region [사용안함]상대적크기 그래픽객체를 anchor에 할당후에 설정해야 합니다.(먼저 등록시 오류발생)
            bool useRelativeSize = false;
            if (useRelativeSize)
            {
                Wp14.RelativeWidth relativeWidth4 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
                Wp14.PercentageWidth percentageWidth4 = new Wp14.PercentageWidth();
                percentageWidth4.Text = "0";
                relativeWidth4.Append(percentageWidth4);

                Wp14.RelativeHeight relativeHeight4 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
                Wp14.PercentageHeight percentageHeight4 = new Wp14.PercentageHeight();
                percentageHeight4.Text = "0";
                relativeHeight4.Append(percentageHeight4);

                anchor.Append(relativeWidth4);
                anchor.Append(relativeHeight4);
            }
            #endregion

            drawing4.Append(anchor);

            alternateContentChoice.Append(drawing4);

            alternateContent.Append(alternateContentChoice);

            return alternateContent;
        }

        private A.Graphic SetText(OfficeModel model)
        {
            var graphic = new A.Graphic();
            graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.GraphicData graphicData4 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape = new Wps.WordprocessingShape();

            #region [뭔지모름] NonVisual DrawingShapeProperties 클래스를 정의합니다. 이 클래스는 Office2010 이상에서 사용할 수 있습니다
            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };
            nonVisualDrawingShapeProperties3.Append(shapeLocks3);
            wordprocessingShape.Append(nonVisualDrawingShapeProperties3);
            #endregion           

            wordprocessingShape.Append(GetShapeProperties(model));

            wordprocessingShape.Append(GetTextBoxInfo2(model));

            wordprocessingShape.Append(GetTextBodyProperties(model));

            graphicData4.Append(wordprocessingShape);

            graphic.Append(graphicData4);

            return graphic;
        }
        /// <summary>
        /// 배경색, 내부편집 컨트롤 Transform, 테두리
        /// </summary>
        private ShapeProperties GetShapeProperties(OfficeModel model)
        {
            var shapeProperties = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            #region 내부 편집 컨트롤 
            shapeProperties.Append(Common.Instance.GetDrawingTransfrom2D(0, 0, model.Rect.Width, model.Rect.Height));
            #endregion

            #region [뭔지모름] 이 요소는 사용자 정의 기하학적 모양 대신 사전 설정된 기하학적 모양을 사용해야 하는 경우를 지정합니다.
            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();
            presetGeometry4.Append(adjustValueList4);
            shapeProperties.Append(presetGeometry4);
            #endregion

            #region 배경색
            if (model.ShapeStyle.UseFill == false)
            {
                shapeProperties.Append(new A.NoFill());
            }
            else
            {
                shapeProperties.Append(Common.Instance.GenerateSolidFill(model.ShapeStyle.FillColor));
            }
            #endregion

            #region 테두리
            if (model.ShapeStyle.UseOutline && model.ShapeStyle.OutlineWeight > 0)
            {
                var outline = Common.Instance.GetDrawingOutline(model.ShapeStyle.OutlineWeight, model.ShapeStyle.OutlineColor);
                shapeProperties.Append(outline);
            }
            #endregion
            return shapeProperties;
        }

        /// <summary>
        /// 가로정렬, 텍스트, 폰트속성
        /// </summary>
        private TextBoxInfo2 GetTextBoxInfo2(OfficeModel model)
        {
            var textBoxInfo = new Wps.TextBoxInfo2();
            TextBoxContent textBoxContent = new TextBoxContent();
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            #region 공백없음 속성
            paragraphProperties.Append(new ParagraphStyleId() { Val = "a3" });
            #endregion
          
            #region 줄간격
            if (model.Paragraph.LineSpacing > 0)
            {
                var lineSpace = model.Paragraph.LineSpacing * 20f;
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = lineSpace.ToString(), LineRule = LineSpacingRuleValues.Exact };
                paragraphProperties.Append(spacingBetweenLines1);
            }
            #endregion


            #region 줄간격
            var spacingBetweenLines = new SpacingBetweenLines() { After = "120", AfterLines = 50, Line = "240", LineRule = LineSpacingRuleValues.Auto };
            paragraphProperties.Append(spacingBetweenLines);
            #endregion

            #region 가로정렬
            paragraphProperties.Append(new Justification() { Val = Common.Instance.GetWordprocessingJustification(model.Paragraph.AlignmentHorizontal) });
            #endregion

            //#region 단락속성
            //ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            //RunFonts runFonts103 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            //paragraphMarkRunProperties.Append(runFonts103);
            //paragraphProperties.Append(paragraphMarkRunProperties);
            //#endregion

            Run run105 = new Run();
            Text text87 = new Text();
            text87.Text = model.Text;

            #region 폰트속성
            var runProperties = Common.Instance.GetWordRunProperty(model.Font);
            run105.Append(runProperties);
            run105.Append(text87);
            #endregion



            paragraph.Append(paragraphProperties);
            paragraph.Append(run105);
            textBoxContent.Append(paragraph);
            textBoxInfo.Append(textBoxContent);
            return textBoxInfo;
        }

        /// <summary>
        /// 여백, 글자방향, 세로정렬, 등
        /// </summary>
        private TextBodyProperties GetTextBodyProperties(OfficeModel model)
        {
            Wps.TextBodyProperties textBodyProperties = new Wps.TextBodyProperties()
            {
                Rotation = 0,
                Vertical = Common.Instance.GetDrawingTextVertical(model.Paragraph.TextDirection),
                Wrap = A.TextWrappingValues.Square, //이걸 적용해야 지정한 사이즈에 딱맞게 생성된다. ppt는 없는듯한데..
                LeftInset = (int)model.Margin.Left * Common.Instance.EMUPPI,
                TopInset = (int)model.Margin.Top * Common.Instance.EMUPPI,
                RightInset = (int)model.Margin.Right * Common.Instance.EMUPPI,
                //BottomInset = (int)model.Margin.Bottom * Common.Instance.EMUPPI,
                BottomInset = 0,
                //Anchor =Common.Instance.GetDrawingAnchoring(model.Paragraph.AlignmentVertical),
                Anchor = A.TextAnchoringTypeValues.Bottom,

                AnchorCenter = false
            };

            #region 도형을 텍스트크기에 맞춤.
            //A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();
            //textBodyProperties.Append(shapeAutoFit1);
            #endregion

            #region 도형을 텍스트크기에 안맞춤.
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();
            textBodyProperties.Append(noAutoFit1);
            #endregion

            return textBodyProperties;
        }
    }

    internal class AnchorProperty
    {
        private AnchorProperty() { }
        //private static 인스턴스 객체
        private static readonly Lazy<AnchorProperty> _instance = new Lazy<AnchorProperty>(() => new AnchorProperty());
        //public static 의 객체반환 함수
        public static AnchorProperty Instance { get { return _instance.Value; } }
        /// <summary>
        /// 특별히 구현은 없음.
        /// 이 요소는 simplePos 속성이 앵커 요소(§20.4.2.3)에 지정된 경우 DrawingML 객체가 페이지의 왼쪽 상단 가장자리를 기준으로 배치되는 좌표를 지정합니다.
        /// </summary>
        internal SimplePosition GetSimplePosition()
        {
            return new SimplePosition() { X = 0L, Y = 0L };
        }

        internal PositionOffset GetPositionOffset(int value)
        {
            PositionOffset positionOffset = new PositionOffset
            {
                Text = (value * Common.Instance.EMUPPI).ToString()
            };
            return positionOffset;
        }

        internal Extent GetExtent(int width, int height)
        {
            return new Extent() { Cx = width * Common.Instance.EMUPPI, Cy = height * Common.Instance.EMUPPI };
        }

        /// <summary>
        /// 도형 효과(현재 없음) //반사 및/또는 그림자 효과
        /// </summary>
        internal EffectExtent GetEffectExtent()
        {
            return new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
        }
    }

}
