using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.DefaultAttributes;
using SimpleOfficeCreator.Stardard.Modules.GeneratedCode;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System.Collections.Generic;
using System.IO;

namespace SimpleOfficeCreator.Stardard.Modules
{
    public class Word
    {
        readonly WordprocessingDocument document;
        readonly MainDocumentPart mainDocumentPart;
        readonly int WORD_RATIO = 15;
        Body body;
        string password = string.Empty;



        public Word(MemoryStream stream)
        {
            document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            mainDocumentPart = document.AddMainDocumentPart();
        }

        public void Save()
        {
            //문서의 모든 컨트롤에 스타일이 적용되는데, 아무것도 지정안하면 기본스타일로 지정되어서 사이즈가 다르게 나온다.
            //그래서 기본 스타일을 추가하고 "간격 없음(a3)" 스타일을 적용한다.
            StyleDefinitionsPart styleDefinitionsPart = this.mainDocumentPart.AddNewPart<StyleDefinitionsPart>("rIdstyle");
            WordBase.Instance.CreateStyleDefinitionsPart(styleDefinitionsPart);

            //필수
            DocumentSettingsPart documentSettingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>("rIdDocumentSettings");
            WordBase.Instance.GenerateDocumentSettingsPart1Content(documentSettingsPart);

            //필수
            EndnotesPart endnotesPart1 = mainDocumentPart.AddNewPart<EndnotesPart>("rIdEndNotes");
            WordBase.Instance.GenerateEndnotesPart1Content(endnotesPart1);

            //필수
            FootnotesPart footnotesPart1 = mainDocumentPart.AddNewPart<FootnotesPart>("rIdFootNotes");
            WordBase.Instance.GenerateFootnotesPart1Content(footnotesPart1);

            //선택
            FontTablePart fontTablePart1 = mainDocumentPart.AddNewPart<FontTablePart>("rIdFontTable");
            WordBase.Instance.GenerateFontTablePartContent(fontTablePart1);

            //선택
            WebSettingsPart webSettingsPart1 = mainDocumentPart.AddNewPart<WebSettingsPart>("rIdWebSettings");
            WordBase.Instance.GenerateWebSettingsPartContent(webSettingsPart1);

            //암호를 지정할수 있지만 파일열때 암호의 경우는 Office 프로그램을 통하여 설정하기 때문에 적용불가.
            //서버모드라면 관련 DLL 설치후 가능할지도..
            //다른 외부 모듈 보니까 파일열어서 적용.
            if (string.IsNullOrEmpty(this.password) == false)
            {
                OpenXMLTools.Tools.ApplyDocumentProtection(document, this.password);
            }

            this.mainDocumentPart.Document.Save();
            this.document.Dispose();
        }

        public void Initialize(int width, int height, string password = "")
        {
            Document document = WordBase.Instance.GenerateDocument();
            this.mainDocumentPart.Document = document;
            this.body = this.mainDocumentPart.Document.Body;

            this.password = password;
        }

        public void ConvertPerPage(int page, List<OfficeModel> models)
        {
            //이미지를 추가한다. 
            Common.Instance.GenerateImagePart(models, this.mainDocumentPart);

            //테이블
            List<OfficeModel> tables = models.FindAll(x => x.Type == Type.Table);
            foreach (OfficeModel table in tables)
            {             
                this.body.Append(SocWordTable.Instance.Generate(table));
            }

            //기타..
            //List<OfficeModel> textboxs = models.FindAll(x => x.Type == Type.TextBox || x.Type == Type.Picture || x.Type == Type.Shape);
            List<OfficeModel> textboxs = models.FindAll(x => x.Type != Type.Table && x.Type != Type.Paper);
            this.body.Append(SocParagraph.Instance.Generate(textboxs));

            //this.body.Append(SocParagraph.Instance.Test2());
            OfficeModel report = models.Find(x => x.Type == Type.Paper);
            //워드 좌측 상단에 줄자가 있고, 여백설정이 가능하다. 
            SectionProperties sectionProperties = new DocumentFormat.OpenXml.Wordprocessing.SectionProperties();
            sectionProperties.Append(용지사이즈설정(report.PaperInfo.Width, report.PaperInfo.Height, report.PaperInfo.IsLandscape));

            //PPT 때문에 절대값으로 다 바꿨는데.. 워드는 여백이 따로 있네?? 제길.
            sectionProperties.Append(용지여백설정(0, 0, 0, 0));
            //sectionProperties.Append(용지여백설정((int)report.Margin.Left, (int)report.Margin.Top, (int)report.Margin.Right, (int)report.Margin.Bottom));
            this.body.Append(sectionProperties);

        }

        private PageSize 용지사이즈설정(int width, int height, bool isLandscape)
        {
            //Twip to Pixel
            PageSize pSize = new PageSize();

            if (isLandscape == false)
            {
                pSize.Orient = PageOrientationValues.Landscape;
                pSize.Width = (uint)width * (uint)WORD_RATIO;
                pSize.Height = (uint)height * (uint)WORD_RATIO;

            }
            else
            {
                pSize.Orient = PageOrientationValues.Portrait;
                pSize.Width = (uint)(width * WORD_RATIO);
                pSize.Height = (uint)(height * WORD_RATIO);
            }
            return pSize;
        }

        private PageMargin 용지여백설정(int left, int top, int right, int bottom)
        {
            PageMargin pageMargin = new PageMargin()
            {
                Left = (uint)(left * WORD_RATIO),
                Top = (int)(top * WORD_RATIO),
                Right = (uint)(right * WORD_RATIO),
                Bottom = (int)(bottom * WORD_RATIO),
                Header = 0U,
                Footer = 0U,
                Gutter = 0U

            };
            return pageMargin;
        }

        private Paragraph 다음페이지로()
        {
            Paragraph p = new Paragraph();
            Run run1 = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };

            run1.Append(break1);
            p.Append(run1);
            return p;
        }


        //todo : 인쇄 배포해야함. PDF 대용량 문제
    }
}
