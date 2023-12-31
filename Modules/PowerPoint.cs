﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Standard.Modules.DefaultCreator;
using SimpleOfficeCreator.Standard.Modules.GeneratedCode;
using SimpleOfficeCreator.Standard.Modules.Model;
using System.Collections.Generic;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
namespace SimpleOfficeCreator.Standard.Modules
{
    /// <summary>
    /// 제약사항
    /// TableCell의 경우 Wrap Text가 자동적용이다. 
    /// 일반 텍스트 상자로 만들경우에는 조절 가능하지만, 일괄 테두리만 적용 가능하다. 
    /// PPT의 용지 프리셋은 실제 사이즈 반영하지 않는다. 너비 높이 제대로 설정하자. 
    /// </summary>
    public class PowerPoint
    {
        PresentationDocument document;
        PresentationPart presentation;

        PPTBase pptDefualt;

        List<string> relationshipIdList = new List<string>();
        public void Save()
        {
            pptDefualt.GeneratePresentationPartContent(presentation, this.relationshipIdList);

            document.Save();
            document.Dispose();
        }
        /// <summary>
        /// 기본스타일을 생성한다.
        /// </summary>
        public PowerPoint(MemoryStream stream, int width = 794, int height = 1123)
        {
            document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true);
            presentation = document.AddPresentationPart();

            pptDefualt = new PPTBase()
            {
                Width = width * Common.Instance.EMUPPI,
                Height = height * Common.Instance.EMUPPI
            };
        }


        public void ConvertPerPage(int page, List<OfficeModel> models)
        {
            //페이지 별로 슬레이드 아이디를 만들어준다. 
            string slideId = "slideId" + (page + 1000);
            relationshipIdList.Add(slideId);
            SlidePart slidePart = this.presentation.AddNewPart<SlidePart>(slideId);

            //이미지를 추가한다. 
            Common.Instance.GenerateImagePart(models, slidePart);

            //기본 레이아웃을 지정한다. 
            if (page == 1)
            {
                pptDefualt.GenerateDefaultSliderPart(slidePart, this.presentation);
            }
            else
            {
                //슬라이드에는 꼭 레이아웃이 지정되어야 문서를 열때 복구 다이얼로그가 생성되지 않는다. 
                pptDefualt.GenerateAddSliderLayoutPart(slidePart);
            }

            //내용을 생성한다. 
            GenerateSlidePartContent(slidePart, models);
        }


        private void GenerateSlidePartContent(SlidePart slidePart1, List<OfficeModel> models)
        {
            #region Default Pre
            var slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            #endregion


            var commonSlideData1 = new CommonSlideData();

            var shapeTree1 = new ShapeTree();
            shapeTree1.Append(GetNonVisualGroupShapeProperties());
            shapeTree1.Append(GetGroupShapeProperties());
            foreach (OfficeModel model in models)
            {
                switch (model.Type)
                {
                    case Model.Type.TextBox:
                        shapeTree1.Append(SocShape.Instance.GenerateTextBox(model));
                        break;
                    case Model.Type.Shape:
                        shapeTree1.Append(SocShape.Instance.GenerateShape(model));
                        break;
                    case Model.Type.Table:
                        shapeTree1.Append(SocPowerpointTable.Instance.Generate(model));
                        break;
                    case Model.Type.Picture:
                    case Model.Type.TableImageCell:
                        shapeTree1.Append(SocShape.Instance.GeneratePicture(model));
                        break;
                }
            }


            commonSlideData1.Append(shapeTree1);
            slide1.Append(commonSlideData1);
            slide1.Append(GenerateColorMapOverride());
            slidePart1.Slide = slide1;

            ColorMapOverride GenerateColorMapOverride()
            {
                ColorMapOverride colorMapOverride1 = new ColorMapOverride();
                A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

                colorMapOverride1.Append(masterColorMapping1);
                return colorMapOverride1;
            }

            NonVisualGroupShapeProperties GetNonVisualGroupShapeProperties()
            {
                NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
                NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
                NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
                ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

                nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
                nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
                nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);
                return nonVisualGroupShapeProperties1;
            }
            GroupShapeProperties GetGroupShapeProperties()
            {
                GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

                A.TransformGroup transformGroup1 = new A.TransformGroup();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 0L, Y = 0L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);
                return groupShapeProperties1;
            }

        }

        //private GraphicFrame GenerateGraphicFrame(OfficeModel model)
        //{
        //    GraphicFrame graphicFrame = new GraphicFrame();
        //    graphicFrame.Append(StaticCode.GenerateNonVisualGraphicFrameProperties("표"));

        //    Transform transform1 = SocPowerpointTable.Instance.Transform(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height);
        //    graphicFrame.Append(transform1);

        //    A.Graphic graphic1 = SocPowerpointTable.Instance.Graphic(model);
        //    //A.Graphic graphic1 = Graphic(model);
        //    graphicFrame.Append(graphic1);

        //    return graphicFrame;
        //}

    }
}
