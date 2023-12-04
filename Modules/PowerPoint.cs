using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SimpleOfficeCreator.Stardard.Modules.DefaultCreator;
using SimpleOfficeCreator.Stardard.Modules.GeneratedCode;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using System.Collections.Generic;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
namespace SimpleOfficeCreator.Stardard.Modules
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

        public const string IMAGE_NOIMAGE = @"R0lGODlhiwCLAPcAAIFuXOTk5MC3rv7+/ufn56GShebm5t/b1urq6ujo6P39/e/v7+vr6/r6+vz8/Ozs7Onp6fn5+fv7+/j4+PHx8e7u7uPj4/X19fT09O3t7fLy8vDw8Pb29vPz8/f399DIwvf29ZmJe+/t64l3ZpGAcLClmambj9fSzMi/t7iupLOqoczHw+fk4JqMfrOpoMvGwqebkL+4sY19bdXQzM3JxLmwqI59bdLPy5SEdtjW1NzY09DMx7mxqbmwp5qLfod1ZcbAuq2imN/e3KCShLKpoNjV06CTh9zX0ruyqfTz8pyOgOXl5f///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwAAAAAiwCLAAAI/wCZCBxIsKDBgwgTKlzIsKHDhxAjSpxIsaLFixgzatzIsaNHhQMaeOiwocIDBBAIqFSZAAGDDAsoXIjgYMDHmzgfOuCw4UECA0uCCh1KlCgBBBU6NLCZsylOBRMoPCBQtKrVqxAqXJDgtGvGCBsSXB1LtqqBBxwUeF37UEIHBmWJBphLt25coQkWTGDKtu/AnRWolrVrIAEDlwwSI0hAAKjdsgg0LPW7NsICwWMDWAhgAEKGDhwiKLDJl8mA0xImXFjAgOrcsgYQYFBLGWeDCkAzB1iCYAGHyRMdTNDwwABdsrJrf8SQ2+prBhT2cmxwIUOC12MrcFX+9YFuCBoiNP9VwCHDEuxVE3QozT2iAgrN5VpYwmA22wmXN++uymBC+4gDXICAcwYwsIF/tTmAgXXoDWXAAtv9t5AD3lU1FwPiSciEAxsYMF96F2iYkAdiFTUXAhewJ2EDC5xn1QK0iTiQBvEJFQABHcQo40ARPLBfUQw0sKMCLZoYAAMR7mgQBS4WlQCCGlZgoQEbOKCkQh4g0GBQCSSpXI8mEqCBl1ceNIAHPhYlpor3lThUAAhkWGZDA9BY1QJsdjUBZkJZgICOczKkAJ9D4enXniYiGehDCqT5Yp43NeCmjTAu6tAAFXx4lQZsVWjjEhRY+tACP15FgJBOKSDlm3GK2pAEjpb/lQAHTq36ppyuJuQAApoShQEDpS5BAK4edWBiqLkutECvQ3XAhATXEQWBlR95UGMAFSS70ATGmbgBj90OlQGkFTkw6RIWMACotgQ1QECw6GZLEAfwcsrRAJ4GBSe17Bo0QbRFGVoQBkUZgKpGF8j1Z78HJQxvAPY2HOwD5EKkAARvnsqwQR40+SYE6xJUQbAYaMTkm8hu/Ne7YRJrUARqkilRBPFxJnO/pJpoAJQKZRCsvBbli+63Ks/48BK0vhquUDxLRPCb6hZtmgYWEpC0QwkTtfBEDdRMAL8q+2xiAgc/REGwC1Bkq77OSp2ziVdbDHBQBkvU9ccha3u2hUA7//2zREXaGGLRMFsYNUUDaDnU1xA5wCecFVu668MIgD3RBcF6ABEHRBngsrYRQPBwpRgBKy5Eax9ZtLsP930RBb167pAEfFpANMPQPgxB5A8VPlTKCxn7ZtPJTm5V3BnFukQCeRcEl40M8H4lvvCe1zZH9BKl+UINyFUyw5hbCPxGF5fq+kEb4M1wBzWeZ8Dgy5VqwM0CKTApxAxrcDTyHTlO1PUHmUBRPmepp1XlezkZ2VAeoJCT6QsB0pMR5+BlgQx0JXtCYRxCnqev8YmKZg/TmFP8N5TtGYSE+iLeogZgOrkQQIU3UZ4HBcK5oUCweELTl6LWEj6h3NAg6bPR7f9E5QAGMOs8leuLBOIzv4PkKwD8mxP1TAQB+uVEcUIx4UAGMKm6iWoAb3uTF/0SxiUAUCB309fuXLWB6tlOORgo1RAF0jEbna9M+ntR85ziu6CkrSDCY5uoslaVM/qFi6crSBD1pUWCHEAAAviAQSApyYKcoAQFKEAJPgCChYiAkgJ5JCQFcICBTCAGLlCBKnPQQSaAYJQiUAgIPoDJApgABZ0kyAdGyctRltIgAxCdUBhgkLWNsSACAIAyfzkQZRaAICwIgTKnCYARnEAhB3CmQJI5zRIIJAIEkAE1V3CeP35gmilIyAlGQE1ljqCSAilAO9spAIRgkTcGyZcB8sb/TQCEoCDaFIgI2DlPZV4TIdkEwDOZ0E9/MkFSRWjnCgJAMYGYYJojQMg5C2rQgciTowCopxN/BAGDDEhfCchTQ1FAkIAy4aPvDKU0q5lLgyR0oQ0FQBJEB4R2vmAJqAJBOw9KEBFM0wQsEMgJSKBMEnhUmSg4gFSneoBYHkRsXDIIxtSoUmqOoKYBTegIkkqQmYrUpgHl5kd3sBsVKLMFyqQB/DbKVAB4syAlUOZdBzLQjr50mRAJYwIMghnVISSn6RRIQPMKgMQ6sqkITasyU6BMIgQFrgAIgjKPQJCLAmCjGSWIUJVZ04GgIAUosOpHmdmQDWiKAISFXkK4WVcA/5A1oDNlbTNJe5CbblOZG21BUJRpAxcAVqDuZIJnicoE3zJktRCBXQZjq8PZKlMAM11oQKeJEOiiVaG/BcAMxPmDJdxAmTBwKwCYiQK9MmGjJiAIN8+qkI9yNCHSDQpsC1JYYh72ugn9LBO2q8zuHrcgzuXmC4ygzBvwQJk1QMKBZ3rN0QKgpvMlSCY3XIBK2reg+H0tdc/j34NkmLEkGK12C3wQ7yJYsgB4gXpjAANlzoCbvzRqNQfiWXhmeLfUFOlHQ8DhDSfEtdMtiJsg99+QMqGvIQ0oQa1akLqS9cXgZUIPlLmCFaDXBsocAI4F0l4AkGCUH42vQODbUnrG8//ADQnjfgly0vOktMkizelCPwpPvnK3t9r0gHEBsAIhKBMHyvznmJlQ24LmUsdOneR137xe1P1osM5bXFedLJBGL7TM/0SmMtX83QJAa9DkRDQ6GQpYFoAUqgOpa5/DK2Q4MwSrSyhpQcwjlCaaeNKhpOZCQUDQvb53mrodSEKHoCX1krPG05TkmCkLABNQ9ZGJHghoWXsAs1I62QppIT4LErigwPDHlM4yq5uaAgHYl9TfVUJbubyEGlAzqWOuK3NdOU0qz7TakPw3p4dc5Ew6liCJ+1GJB+LA88BP1JxGrkuZwNh5hqC0WPZBUJwdABpQM7wHcPWODeJZlgoEBAL/b2cKcvnheS60IFz8EQMLUkN9zXEg6A7vy5Uq8DMvRAdv3bgyX2AADkwTp8uktrG1nW2CoKDRIygBlf8K0p0PZIlDOV8fDasREEh16gnJwBF3E0WNiODrGsFgUG6OQmFZzisG9F6yNNCriG2xzub2ywRNZHdRKe/hA+G1IO/DMrncfFGDIsrnAnmemXvFA4UnyuEXVUe8hKx7RLEiR4xFwTtaqo1ZR0gw34TAnFReLtNiVzBLVfqCpM6CTVn92DRfpj4GpWwFIaR+396RMupLhOzKo1DWiBCs26jsJnvYzjZ2zyX8MSFPHNdH3mMVAjQyWQIkCgwHUvMUeqRRR8z7/8aUt3CE2O9Nz9cIpqoXgBnmKo1C6TtCyu12jghfLunv1956jXuEZL9PhlQRtPMwFbUx9lMqjrcQifMmsIcRDiBMRJEueyQqcUQUgKcQVDMUCcB7EdEA4iYuHJgsYjcUIPMQmCdEF8EtoyM1z1IjraeAeMcZ/WeCkWcjGlQ0vkdACRF35UQR5XM8LNg1pfJDjLJV+iI7EnF/b3KBDKNAQ/GCDaF7jScRF7A0NhKA/fJ/PhRBCLc25+F+CJGBLoR82pI7i6ODDNF2QXF9CLF/YrR92jJFv3MRjHcexKcQdWgjUMgwFUiCE8gQC/gmbEgQPNgnRFg0B2iBGcEtUAMpS/+kO7SXK30oFA2IEfSHP2aCazZCNiwoEOZSKgQQiQ6hhqBiEBLwgfryAKIoKg0AgULBhBYhhedReg/4MIfDgj84FNHjEZdoANujAKhoZyGoepooLDOIEQqAd+cRigMwgiaChQwDeop4E7SjML5nc50oEB0AL5O3Ef8iF9WzBBvAhYvSfUIhMDlRjWWxh/lTFfKHE3koF9CoetKoi2wBRmMBhuziAM54hn2Bj1SUja0ILy9EGQApRrCoLbWoJnAYexpAKEFRAWi4KAoyN0LxABPZFe5iIkuQAccYKBjwLvUiIpA3JQmpJLCyJX60I6FjIaDyh/9BIuFIAeTYFfzoHAn/gAE16RUNYB7VQwAn2R510j7nASdkSBks4iFWESSW0gA5pC/0QQEZOR7VYRw/mSKuMgAYcC42shtoAZMd0QAUcB3h2JEfOScS4IXgCAH2oScZYJVYEZSLMgHKCI4JUAEcsIoQMQARoAEDUpbCQgHD6CpEQpRdeSMPsBU7WRBQsQB/CZi8MZW5EgGBMRic0RsYIBq8cxoNwAEaYB1FCRltmY2e6AGCpxu7QQAQwAAPUAELsAEUEJsbsAAVkAGH8ROPMRayEQGLmSxgYYRkURdzYQHEWZzCeRdncQFgmY0K4AG4cRfQSRYMIBmkORFuMRXRmZ28oRfV6YAeQAEIYJjaK3mXmdmb3Vl8I0EBC2CbKJEAjLESLAEBLtGaG6ABM7Gc55mf+rmf/CkjAQEAOw==";
        int EMUPPI = 9525;
        PPTDefault pptDefualt;

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
        public PowerPoint(MemoryStream stream)
        {
            document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true);
            presentation = document.AddPresentationPart();
        }

        public void Initialize(int width, int height, int emuppi)
        {
            EMUPPI = emuppi;
            pptDefualt = new PPTDefault()
            {
                Width = EMUPPI * width,
                Height = EMUPPI * height
            };
            GenerateTableCell.Instance.EMUPPI = emuppi;
            GenerateTextBox.Instance.EMUPPI = emuppi;
        }

        public void ConvertPerPage(int page, List<OfficeModel> models)
        {
            //페이지 별로 슬레이드 아이디를 만들어준다. 
            string slideId = "slideId" + (page + 1000);
            relationshipIdList.Add(slideId);
            SlidePart slidePart = this.presentation.AddNewPart<SlidePart>(slideId);

            //이미지를 추가한다. 
            var pictures = models.FindAll(x => x.Type == Model.Type.Picture);
            foreach (var picture in pictures)
            {
                GenerateImagePart(slidePart, picture.Text, picture.UID);
            }

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
            shapeTree1.Append(StaticCode.GenerateNonVisualGroupShapeProperties());
            shapeTree1.Append(StaticCode.GenerateGroupShapeProperties());
            foreach (OfficeModel model in models)
            {
                switch (model.Type)
                {
                    case Model.Type.TextBox:
                        shapeTree1.Append(GenerateTextBox.Instance.Generate(model));
                        break;
                    case Model.Type.Shape:
                        shapeTree1.Append(GenerateTextBox.Instance.GenerateShape(model));
                        break;
                    case Model.Type.Table:
                        shapeTree1.Append(GenerateGraphicFrame(model));
                        break;
                    case Model.Type.Picture:
                        shapeTree1.Append(GenerateTextBox.Instance.GeneratePicture(model));
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
        }


        private GraphicFrame GenerateGraphicFrame(OfficeModel model)
        {
            GraphicFrame graphicFrame = new GraphicFrame();
            graphicFrame.Append(StaticCode.GenerateNonVisualGraphicFrameProperties("표"));

            Transform transform1 = GenerateTableCell.Instance.Transform(model.Rect.X, model.Rect.Y, model.Rect.Width, model.Rect.Height);
            graphicFrame.Append(transform1);

            A.Graphic graphic1 = GenerateTableCell.Instance.Graphic(model);
            //A.Graphic graphic1 = Graphic(model);
            graphicFrame.Append(graphic1);

            return graphicFrame;
        }

        #region Binary Data
        private void GenerateImagePart(SlidePart slidePart, string base64, string id)
        {
            ImagePart imagePart = slidePart.AddNewPart<ImagePart>("image/png", id);

            if (string.IsNullOrEmpty(base64) || base64.Contains("http"))
                base64 = IMAGE_NOIMAGE;

            System.IO.Stream data = GetBinaryDataStream(base64);
            imagePart.FeedData(data);
            data.Close();

        }
        private readonly string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCADAAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAorAvvHWi6bqFxZXN1JFc2728cqm2lwnnsyxMW242MyMu/O0FSCQav6Prtlr8M0tjK0iwytBKrxtG8bjBKsrAEHBB5HIII4IoA0KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA84ufBXg/xd408QwyG5l1qN7C51CNomjG1DI1sVcoAw3CQh1YspGAy4ArrfCvhHT/B9ndW+nq4W6uXu5S7D5pGwCcAADoOgGTljlmYl9h4V03TfEGpa3BAy6nqMcUVzM0rtuWPdsUAnAA3t0A6/StegAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/9k=";

        private Stream GetBinaryDataStream(string base64String)
        {
            if (base64String.Contains("http"))
            {
                Logger.Instance.Write("이미지 경로에 http가 포함됨");
                return new MemoryStream(Convert.FromBase64String(thumbnailPart1Data));
            }


            return new MemoryStream(Convert.FromBase64String(base64String));
        }

        #endregion
    }
}
