using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using SimpleOfficeCreator.Stardard.Modules.Model;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab;
using System;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class GenerateTextBox
    {
        private GenerateTextBox() { }
        //private static 인스턴스 객체
        private static readonly Lazy<GenerateTextBox> _instance = new Lazy<GenerateTextBox>(() => new GenerateTextBox());
        //public static 의 객체반환 함수
        public static GenerateTextBox Instance { get { return _instance.Value; } }

        public int EMUPPI { get; set; } = 0;

        /// <summary>
        /// PPT에서 일반 텍스트 label은 Shape 이라고 간주한다.
        /// </summary>
        public Shape Generate(OfficeModel model)
        {
            int fontSize = (int)model.Font.Size * 100;
            string fontFace = model.Font.Name;
            A.TextAlignmentTypeValues textAlignment = model.Paragraph.AlignmentHorizontal;
            A.TextAnchoringTypeValues textAnchoring = model.Paragraph.AlignmentVertical;

            bool bold = false;
            bool italic = false;
            A.TextUnderlineValues underlineValue = A.TextUnderlineValues.None;
            if (model.Font.UnderLine)
                underlineValue = A.TextUnderlineValues.Single;

            A.TextStrikeValues strikeValues = A.TextStrikeValues.NoStrike;
            if (model.Font.Strike)
                strikeValues =  A.TextStrikeValues.SingleStrike;

            


            //텍스트 수직정렬(방향)
            A.TextVerticalValues textVerticalValue = A.TextVerticalValues.Horizontal;
            switch (model.Paragraph.TextDirection)
            {
                case TextDirection.Vertical:
                    textVerticalValue = A.TextVerticalValues.EastAsianVetical;
                    break;
                case TextDirection.RotateAllText90:
                    textVerticalValue = A.TextVerticalValues.Vertical;
                    break;
                case TextDirection.RotateAllText270:
                    textVerticalValue = A.TextVerticalValues.Vertical270;
                    break;
                case TextDirection.Stacked:
                    textVerticalValue = A.TextVerticalValues.WordArtVertical;
                    break;
            }
            TextBody textBody1 = new TextBody(); 
            
            SetBodyProperty();
            textBody1.Append(new A.ListStyle());

            //아래값은 텍스트 속성의 자동 크기조정 같은 기능인데, PPT 내부에서 계산되는 값이므로, 변환시에 속성을 셋팅으로 설정이 불가능
            //A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit() { FontScale = 70000, LineSpaceReduction = 20000 };
            //변환시 글자 가운데 공백을 많이 넣은경우, 공백 사이즈에 의해서 영역을 벗어나는 케이스가 있다. 

            //1. 단락 속성 : 정렬
            //가로 정렬 적용(순서 주위 할 것)
            A.Paragraph paragraph1 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties()
            {
                Alignment = textAlignment
            };
            paragraph1.Append(paragraphProperties2);

            A.Run run1 = new A.Run();
            //2. 텍스트 속성 : 폰트, 컬러, 내용, Bold 등
            SetTextProperty();
            paragraph1.Append(run1);  
            textBody1.Append(paragraph1);

            Shape shape1 = new Shape();
            shape1.Append(StaticCode.GenerateNonVisualShapeProperties("textbox"));

            ShapeProperties shapeProperties1 = new ShapeProperties();
            //배경색, 테두리등.
            SetShapeProperty();
            shape1.Append(shapeProperties1);

            shape1.Append(textBody1);
            return shape1;

            void SetBodyProperty()
            {
                //TextWrappingValues.None 상자안에 아이템이 들어갈 필요가 없다.
                //세로 정렬은 여기
                //컨트롤 내부 여백도 여기
                //1cm = 360000 ex)0.2cm = 72000
                A.BodyProperties bodyProperties1 = new A.BodyProperties()
                {
                    Vertical = textVerticalValue,
                    Wrap = A.TextWrappingValues.None,
                    RightToLeftColumns = false,
                    Anchor = textAnchoring,
                    LeftInset = (int)model.Margin.Left * EMUPPI,
                    TopInset = (int)model.Margin.Top * EMUPPI,
                    RightInset = (int)model.Margin.Right * EMUPPI,
                    BottomInset = (int)model.Margin.Bottom * EMUPPI
                };
                A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();
                bodyProperties1.Append(shapeAutoFit1);
                textBody1.Append(bodyProperties1);
            }
            void SetShapeProperty()
            {
                shapeProperties1.Append(Transform2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));

                //도형 타입
                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                presetGeometry1.Append(new A.AdjustValueList());
                shapeProperties1.Append(presetGeometry1);

                #region 배경색
                if (model.ShapeStyle.UseFill == false)
                {
                    shapeProperties1.Append(new A.NoFill());
                }
                else
                {
                    //A.SolidFill backgroundColor = new();
                    //A.RgbColorModelHex rgbBackColor = new A.RgbColorModelHex() { Val = model.ShapeStyle.FillColor };
                    //if (model.ShapeStyle.FillColor == "transparent")
                    //{
                    //    A.Alpha alpha4 = new A.Alpha() { Val = 0 };
                    //    rgbBackColor.Append(alpha4);
                    //}
                    //backgroundColor.Append(rgbBackColor);
                    shapeProperties1.Append(Common.Instance.GenerateSolidFill(model.ShapeStyle.FillColor));
                }
                #endregion

                if (model.ShapeStyle.UseOutline && model.ShapeStyle.OutlineWeight > 0)
                {
                    A.Outline outline1 = new A.Outline()
                    {
                        Width = (int)model.ShapeStyle.OutlineWeight * EMUPPI,
                    };
                    A.SolidFill solidFill2 = Common.Instance.GenerateSolidFill(model.ShapeStyle.OutlineColor);
                    outline1.Append(solidFill2);
                    shapeProperties1.Append(outline1);
                }
            }
            void SetTextProperty()
            {
                //point * 100
                A.RunProperties runProperties1 = new A.RunProperties() { 
                    Language = "en-US", 
                    AlternativeLanguage = "ko-KR", 
                    Dirty = false,
                    FontSize = fontSize,
                    Bold = bold,
                    Italic = italic,
                    Underline = underlineValue,
                    Strike = strikeValues,
                    Spacing = (int)(model.Font.SpacingValue * 100)

                };

                //주의! 컬러값 잘못들어가면 문서깨짐.
                A.SolidFill solidFill1 = Common.Instance.GenerateSolidFill(model.Font.Color);
                A.LatinFont latinFont1 = new A.LatinFont() { Typeface = fontFace };
                A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = fontFace };

                runProperties1.Append(solidFill1);
                runProperties1.Append(latinFont1);
                runProperties1.Append(eastAsianFont1);

                A.Text text1 = new A.Text();
                text1.Text = model.Text;

                run1.Append(runProperties1);
                run1.Append(text1);
            }
        }

        public Shape GenerateShape(OfficeModel model)
        {


            //아래값은 텍스트 속성의 자동 크기조정 같은 기능인데, PPT 내부에서 계산되는 값이므로, 변환시에 속성을 셋팅으로 설정이 불가능
            //A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit() { FontScale = 70000, LineSpaceReduction = 20000 };
            //변환시 글자 가운데 공백을 많이 넣은경우, 공백 사이즈에 의해서 영역을 벗어나는 케이스가 있다. 

            //1. 단락 속성 : 정렬
            //가로 정렬 적용(순서 주위 할 것)

            Shape shape1 = new Shape();
            shape1.Append(StaticCode.GenerateNonVisualShapeProperties("shape"));

            ShapeProperties shapeProperties1 = new ShapeProperties();
            //배경색, 테두리등.
            SetShapeProperty();
            shape1.Append(shapeProperties1);

            return shape1;

            void SetBodyProperty()
            {
                A.BodyProperties bodyProperties1 = new A.BodyProperties()
                {
                    Wrap = A.TextWrappingValues.None,
                    RightToLeftColumns = false
                };
                bodyProperties1.Append(new A.ShapeAutoFit());
            }
            void SetShapeProperty()
            {
                shapeProperties1.Append(Transform2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));

                //도형 타입
                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
                presetGeometry1.Append(new A.AdjustValueList());
                shapeProperties1.Append(presetGeometry1);

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

                if (model.ShapeStyle.UseOutline && model.ShapeStyle.OutlineWeight > 0)
                {
                    A.Outline outline1 = new A.Outline()
                    {                        
                        Width = (int)model.ShapeStyle.OutlineWeight * EMUPPI,
                    };
                    A.SolidFill solidFill2 = Common.Instance.GenerateSolidFill(model.ShapeStyle.OutlineColor);
                    outline1.Append(solidFill2);
                    shapeProperties1.Append(outline1);
                }
            }            
        }
        public A.Transform2D Transform2D(int x, int y, int width, int height)
        {
            A.Transform2D transform1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = x * EMUPPI, Y = y * EMUPPI };
            A.Extents extents1 = new A.Extents() { Cx = width * EMUPPI, Cy = height * EMUPPI };

            transform1.Append(offset1);
            transform1.Append(extents1);
            return transform1;
        }

        public Picture GeneratePicture(OfficeModel model)
        {
            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "그림 12" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties1);

            BlipFill blipFill1 = new BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = model.UID };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            //shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(Transform2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            return picture1;
        }

    }
}
