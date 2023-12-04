using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficeCreator.Stardard.Modules.DefaultAttributes;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules
{
    public class Word
    {
        WordprocessingDocument document;
        MainDocumentPart mainDocument;
        Body body;

        public Word(MemoryStream stream)
        {
            document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            mainDocument = document.AddMainDocumentPart();

            

        }

        public void Initialize(int width, int height)
        {
            WORDDefault WORDDefault = new WORDDefault();

            //문서의 모든 컨트롤에 스타일이 적용되는데, 아무것도 지정안하면 기본스타일로 지정되어서 사이즈가 다르게 나온다.
            //그래서 기본 스타일을 추가하고 "간격 없음(a3)" 스타일을 적용한다.
            StyleDefinitionsPart styleDefinitionsPart1 = mainDocument.AddNewPart<StyleDefinitionsPart>("rId1");
            WORDDefault.CreateStyleDefinitionsPart(styleDefinitionsPart1);

            new Document(new Body()).Save(mainDocument);
            this.body = mainDocument.Document.Body;
        }

        /// <summary>
        /// 사전에 해야 할 작업
        /// 모든 위치를 절대값으로 변경
        /// 
        /// </summary>
        /// <param name="pageObject"></param>
        //public void Convert(List<OfficeObject> pageObject)
       // {
            //목록에서 도형 분리
            //목록에서 이미지 분리

            //목록에서 테이블 분리


            //라벨 표를 먼저 작성한다.


            //도형 
            //이미지 를 추가한다.

            //용지사이즈, 여백, 방향을 설정한다.

            //다음페이지가 있을 경우에 페이지 브레이크를 추가한다.?? 이건 처음에 추가해야할듯.

       // }
        
    }
}
