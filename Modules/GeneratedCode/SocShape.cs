using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class SocShape
    {
        private SocShape() { }
        //private static 인스턴스 객체
        private static readonly Lazy<SocShape> _instance = new Lazy<SocShape>(() => new SocShape());
        //public static 의 객체반환 함수
        public static SocShape Instance { get { return _instance.Value; } }

        public int EMUPPI { get; set; } = 0;

        /// <summary>
        /// PPT에서 일반 텍스트 label은 Shape 이라고 간주한다.
        /// </summary>
        public Shape GenerateTextBox(OfficeModel model)
        {
            //아래값은 텍스트 속성의 자동 크기조정 같은 기능인데, PPT 내부에서 계산되는 값이므로, 변환시에 속성을 셋팅으로 설정이 불가능
            //A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit() { FontScale = 70000, LineSpaceReduction = 20000 };
            //변환시 글자 가운데 공백을 많이 넣은경우, 공백 사이즈에 의해서 영역을 벗어나는 케이스가 있다. 

            //1. 단락 속성 : 정렬
            //가로 정렬 적용(순서 주위 할 것)
            A.Paragraph paragraph1 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties()
            {
                Alignment = Common.Instance.GetDrawingAlignment(model.Paragraph.AlignmentHorizontal)
            };
            paragraph1.Append(paragraphProperties2);

            #region Run
            //2. 텍스트 속성 : 폰트, 컬러, 내용, Bold 등
            A.Run run1 = new A.Run();
            A.RunProperties runProperties = Common.Instance.GetDrawingRunProperty(model.Font);
            run1.Append(runProperties);
            run1.Append(Common.Instance.GetDrawingRunText(model));
            #endregion
            paragraph1.Append(run1);

            TextBody textBody1 = new TextBody();
            textBody1.Append(GetBodyProperty(model));
            textBody1.Append(new A.ListStyle());
            textBody1.Append(paragraph1);

            Shape shape1 = new Shape();
            shape1.Append(GetNonVisualShapeProperties("textbox"));
            //배경색, 테두리등.
            shape1.Append(GetShapeProperty(model));
            shape1.Append(textBody1);
            return shape1;


        }

        public Shape GenerateShape(OfficeModel model)
        {
            //아래값은 텍스트 속성의 자동 크기조정 같은 기능인데, PPT 내부에서 계산되는 값이므로, 변환시에 속성을 셋팅으로 설정이 불가능
            //A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit() { FontScale = 70000, LineSpaceReduction = 20000 };
            //변환시 글자 가운데 공백을 많이 넣은경우, 공백 사이즈에 의해서 영역을 벗어나는 케이스가 있다. 

            //1. 단락 속성 : 정렬
            //가로 정렬 적용(순서 주위 할 것)

            Shape shape1 = new Shape();
            shape1.Append(GetNonVisualShapeProperties("shape"));

            //배경색, 테두리등.
            shape1.Append(GetShapeProperty(model));

            return shape1;

            //void SetBodyProperty()
            //{
            //    A.BodyProperties bodyProperties1 = new A.BodyProperties()
            //    {
            //        Wrap = A.TextWrappingValues.None,
            //        RightToLeftColumns = false
            //    };
            //    bodyProperties1.Append(new A.ShapeAutoFit());
            //}

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

            shapeProperties1.Append(Common.Instance.GetDrawingTransfrom2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();
            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            return picture1;
        }


        private A.BodyProperties GetBodyProperty(OfficeModel model)
        {
            //TextWrappingValues.None 상자안에 아이템이 들어갈 필요가 없다.
            //세로 정렬은 여기
            //컨트롤 내부 여백도 여기
            //1cm = 360000 ex)0.2cm = 72000
            A.BodyProperties bodyProperties1 = new A.BodyProperties()
            {
                //텍스트 방향
                Vertical = Common.Instance.GetDrawingTextVertical(model.Paragraph.TextDirection),
                Wrap = A.TextWrappingValues.None,
                RightToLeftColumns = false,
                //텍스트 수직정렬
                Anchor = Common.Instance.GetDrawingAnchoring(model.Paragraph.AlignmentVertical),
                LeftInset = (int)model.Margin.Left * EMUPPI,
                TopInset = (int)model.Margin.Top * EMUPPI,
                RightInset = (int)model.Margin.Right * EMUPPI,
                BottomInset = (int)model.Margin.Bottom * EMUPPI
            };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();
            bodyProperties1.Append(shapeAutoFit1);

            return bodyProperties1;
        }
        private ShapeProperties GetShapeProperty(OfficeModel model)
        {
            ShapeProperties shapeProperties1 = new ShapeProperties();

            shapeProperties1.Append(Common.Instance.GetDrawingTransfrom2D(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height));

            //도형 타입
            shapeProperties1.Append(Common.Instance.GetPresetGeometry(model.ShapeStyle.ShapeTypeValue));

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
                A.Outline outline = Common.Instance.GetDrawingOutline(model.ShapeStyle.OutlineWeight, model.ShapeStyle.OutlineColor);
                shapeProperties1.Append(outline);
            }
            #endregion

            return shapeProperties1;
        }

        private NonVisualShapeProperties GetNonVisualShapeProperties(string textboxName)
        {
            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = textboxName };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            return nonVisualShapeProperties1;
        }
    }
}
