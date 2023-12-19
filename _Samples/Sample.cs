using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using SimpleOfficeCreator.Standard.Modules.Model;
using SimpleOfficeCreator.Standard.Modules.Model.Component.HomeTab;
using SimpleOfficeCreator.Standard.Modules.Model.Component.PictureFormatTab;
using SimpleOfficeCreator.Standard.Modules.Model.Component.ShapeFormat;
using SimpleOfficeCreator.Standard.Modules.Model.Component.TableDesignTab;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace SimpleOfficeCreator.Standard._Samples
{
    public class Sample
    {

        public string CreateSingleImageDocument()
        {
            //생성할 문서를 지정한다.
            var officeCreator = new OfficeCreator(OfficeType.Word);

            //이미지를 가져온다.
            string base64String = Utils.Instance.GetWebImage("https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31");

            //테두리 스타일 지정
            var officePictureStyle = new OfficePictureStyle()
            {
                Weight = 3,
                Color = "blue",
                NoOutline = false,
                Dashes = "DashDotDot"
            };

            //이미지 생성
            var model = new OfficeModelCreator().CreatePicture(0, 0, 100, 100, base64String, officePictureStyle);
            
            //모델 목록을 만들고 이미지를 추가한다. 
            var officeModels = new List<OfficeModel>();
            officeModels.Add(model);

            //변환한다. 
            officeCreator.ConvertPage(1, officeModels);

            return officeCreator.Save();
        }

        public string CreateSingleTextBoxDocument()
        {
            //생성할 문서를 지정한다.
            var officeCreator = new OfficeCreator(OfficeType.Word);

            //글꼴을 설정합니다.[옵션]
            var font = new OfficeFont()
            {
                Name = "맑은 고딕",
                Size = 10,
                Bold = false,
                Italic = false,
                UnderLine = false,
                Strike = false,
                CharacterSpacing = 0,
                Color = "black"
            };

            //단락을 설정합니다.[옵션]
            var paragraph = new OfficeParagraph()
            {
                AlignmentHorizontal = TextAlignmentHorizontal.Center,
                AlignmentVertical = TextAlignmentVertical.Center,
                TextDirection = TextDirection.Stacked,
                LineSpacing = 0
            };

            var style = new OfficeShapeStyle()
            {
                UseFill = true,
                FillColor = "yellow",
                UseOutline = true,
                OutlineWeight = 1,
                OutlineColor = "black",
                OutlineDashes = "solid",
                ShapeTypeValue = "rectangle"
            };


            //텍스트 박스 생성
            var model = new OfficeModelCreator().CreateTextBox(50,50, 100, 100, "텍스트 박스 \n테스트",  font, paragraph, style); 

            //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
            var officeModels = new List<OfficeModel>();
            officeModels.Add(model);

            //변환한다. 
            officeCreator.ConvertPage(1, officeModels);

            return officeCreator.Save();
        }

        public string CreatesingleShapeDocument()
        {
            //생성할 문서를 지정한다.
            var officeCreator = new OfficeCreator(OfficeType.Word);

            //테두리 스타일 지정
            var style = new OfficeShapeStyle()
            {
                UseFill = true,
                FillColor = "green",
                UseOutline = true,
                OutlineWeight = 1,
                OutlineColor = "black",
                OutlineDashes = "solid",
                ShapeTypeValue = "rectangle"
            };

            //도형 생성
            var model = new OfficeModelCreator().CreateShape(50, 30, 100, 100, style);

            //모델 목록을 만들고 이미지를 추가한다. 
            var officeModels = new List<OfficeModel>();
            officeModels.Add(model);

            //변환한다. 
            officeCreator.ConvertPage(1, officeModels);

            return officeCreator.Save();
        }


        public string CreateSingleTableDocument()
        {
            //생성할 문서를 지정한다.
            var officeCreator = new OfficeCreator(OfficeType.PowerPoint);
            //모델 목록을 만들고 테이블을 추가한다. 
            var officeModels = new List<OfficeModel>();

            //글꼴을 설정합니다.[옵션]
            var font = new OfficeFont()
            {
                Name = "맑은 고딕",
                Size = 10,
                Bold = false,
                Italic = false,
                UnderLine = false,
                Strike = false,
                CharacterSpacing = 0,
                Color = "black"
            };

            //단락을 설정합니다.[옵션]
            var paragraph = new OfficeParagraph()
            {
                AlignmentHorizontal = TextAlignmentHorizontal.Center,
                AlignmentVertical = TextAlignmentVertical.Center,
                TextDirection = TextDirection.Horizontal,
                LineSpacing = 0
            };
            //테이블 스타일을 설정합니다.[옵션]
            var style = new OfficeTableStyles()
            {
                UseShading = true,
                ShadingColor = "yellow"
            };
            var border = new Modules.Model.Component.TableDesignTab.Border()
            {
                Draw = true, 
                Color = "red",
                Weight = 1,
                Dashes = "solid"
            };
            style.Bottom = border;

            var style2 = new OfficeTableStyles()
            {
                UseShading = true,
                ShadingColor = "transparent"
            };
            var style3 = new OfficeTableStyles()
            {
                UseShading = true,
                ShadingColor = "green"
            };

            #region Case1
            var table1 = new OfficeModelCreator().CreateTable(50, 30, 300, 300, 2, 2);
            officeModels.Add(table1);

            new OfficeModelCreator().CreateTableCell(table1, 0, 0, "PowerPoint" ,1 , 1 , font, paragraph, style);
            new OfficeModelCreator().CreateTableCell(table1, 0, 1, "Word", 1 , 1 , font, paragraph, style);
            new OfficeModelCreator().CreateTableCell(table1, 1, 0, "조금씩", 1 , 1 , font, paragraph, style);
            new OfficeModelCreator().CreateTableCell(table1, 1, 1, "다르다", 1 , 1 , font, paragraph, style);
            #endregion

            #region Case2 셀 병합

            var table2 = new OfficeModelCreator().CreateTable(400, 30, 300, 300, 2, 2);
            officeModels.Add(table2);
            
            new OfficeModelCreator().CreateTableCell(table2, 0, 0, "PowerPoint", 1, 2, null, null, style2);
            new OfficeModelCreator().CreateTableCell(table2, 1, 0, "조금씩", 1, 1, null, null, style2);
            new OfficeModelCreator().CreateTableCell(table2, 1, 1, "다르다", 1, 1, null, null, style2);
            #endregion

            #region Case2 Col & Row 사이즈 지정 및 이미지 추가
            List<int> columnWidth = new List<int>();
            columnWidth.Add(100);
            columnWidth.Add(500);
            List<int> rowHeight = new List<int>();
            rowHeight.Add(100);
            rowHeight.Add(200);

            //테이블 생성
            var table3 = new OfficeModelCreator().CreateTable(50, 350, 600, 300, columnWidth, rowHeight);
            officeModels.Add(table3);
            
            new OfficeModelCreator().CreateTableCell(table3, 0, 0, "사이즈 지정", 1, 1, font, paragraph, style3);
            new OfficeModelCreator().CreateTableCell(table3, 0, 1, "이미지 테이블내 삽입 워드만 가능", 1, 1, font, paragraph, style3);
            new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이미지", 1, 1, font, paragraph, style2);

            //이미지를 테이블내에 삽입하는건 워드만 가능합니다. PPT는 해당 기능이 없이 그냥 지정한 위치에 이미지가 올라가게 됩니다. 
            if(officeCreator.OfficeType == OfficeType.Word)
            {
                string base64String = Utils.Instance.GetWebImage("https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31");
                //이미지 생성
                var image = new OfficeModelCreator().CreatePicture(0, 0, 100, 100, base64String);
                //테이블 내에 넣는 이미지라서 타입을 변경한다. 
                image.Type = Modules.Model.Type.TableImageCell;
                officeModels.Add(image);
                new OfficeModelCreator().CreateTableCell(table3, 1, 1, "", 1, 1, font, paragraph, style2, image.UID);
            }
            else
            {
                new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이 좌표 계산해서 이미지를 추가한다. ", 1, 1, font, paragraph, style2);
            }
            #endregion

            //변환
            officeCreator.ConvertPage(1, officeModels);

            return officeCreator.Save();
        }

        public string CreateCommonProperty()
        {
            //생성할 문서를 지정한다.
            //오피스에는 페이지별로 용지 사이즈 설정기능이 없습니다.
            var officeCreator = new OfficeCreator(OfficeType.Word, 600, 700);

            {
                //여백이 정상인지 확인하기 위해서 테두리 표시
                var style = new OfficeShapeStyle()
                {
                    UseOutline = true
                };

                //텍스트 박스 생성
                var model = new OfficeModelCreator().CreateTextBox(50, 50, 100, 100, "1번페이지", null, null, style);
                model.Margin.Left = 20;
                model.Margin.Top = 50;
                //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
                var officeModels = new List<OfficeModel>();
                officeModels.Add(model);

                //변환한다. 
                officeCreator.ConvertPage(1, officeModels);
            }
            {
                //텍스트 박스 생성
                var model = new OfficeModelCreator().CreateTextBox(50, 50, 100, 100, "2번 페이지");
                //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
                var officeModels = new List<OfficeModel>();
                officeModels.Add(model);

                //변환한다. 
                officeCreator.ConvertPage(2, officeModels);
            }
            return officeCreator.Save();

        }
    }
}
