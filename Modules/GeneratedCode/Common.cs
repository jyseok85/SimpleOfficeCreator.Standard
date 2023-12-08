using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SimpleOfficeCreator.Stardard.Modules.Model;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class Common
    {
        private Common() { }
        //private static 인스턴스 객체
        private static readonly Lazy<Common> _instance = new Lazy<Common>(() => new Common());
        //public static 의 객체반환 함수
        public static Common Instance { get { return _instance.Value; } }

        public List<uint> UniqueId { get; set; } = new List<uint>();
        public int EMUPPI { get; set; } = 9525;

        /// <summary>
        /// 배경색 컨트롤을 생성한다.
        /// </summary>
        /// <returns></returns>
        public Drawing.SolidFill GenerateSolidFill(string color)
        {
            var borderColor = color;
            //투명으로 들어왔다면 그냥 흰색으로 변경한다. 대신 추후 Alpha 컴포넌트를 추가한다. 
            if (borderColor == "transparent" || borderColor == "trasnparent") //오타뭐임??
                borderColor = "FFFFFF";

            Drawing.SolidFill solidFill = new Drawing.SolidFill();
            Drawing.RgbColorModelHex rgbBackColor = new Drawing.RgbColorModelHex() { Val = borderColor };
            if (color == "transparent")
            {
                Drawing.Alpha alpha = new Drawing.Alpha() { Val = 0 };
                rgbBackColor.Append(alpha);
            }
            solidFill.Append(rgbBackColor);

            return solidFill;
        }

        /// <summary>
        /// 오피스에서는 RGB Hex 값이 사용되며, 투명은 NoFill 속성으로 처리된다.
        /// </summary>
        /// <param name="text">컬러명</param>
        /// <returns>RGB Hex 값</returns>
        public string GetOfficeColor(string text)
        {
            if (text == null)
            {
                return "transparent";
            }
            if (text.Contains("#"))
            {
                return text;
            }
            else if (text.ToLower().Equals("transparent"))
            {
                return "transparent";
            }
            else if (text.Contains(","))
            {
                var value = text.Split(',').Select(Int32.Parse).ToList();
                System.Drawing.Color myColor = System.Drawing.Color.FromArgb(value[0], value[1], value[2]);
                string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
                return hex;
            }
            else
            {
                System.Drawing.Color myColor = System.Drawing.Color.FromName(text);
                string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
                return hex;
            }
        }

        public void GenerateImagePart(List<OfficeModel> models, OpenXmlPart openXmlPart)
        {
            //이미지를 추가한다. 
            var pictures = models.FindAll(x => x.Type == Model.Type.Picture);
            foreach (var picture in pictures)
            {
                var base64 = picture.Text;
                var id = picture.UID;
                ImagePart imagePart = openXmlPart.AddNewPart<ImagePart>("image/png", id);

                if (string.IsNullOrEmpty(base64) || base64.Contains("http"))
                {
                    //이미지는 base64 문자열로 입력되어야한다. 만약 http 가 그대로 붙어 있다면 Json저장을 못한 케이스다.
                    base64 = IMAGE_NOIMAGE;
                }

                System.IO.Stream data = GetBinaryDataStream(base64);
                imagePart.FeedData(data);
                data.Close();
            }
        }

        private const string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCADAAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAorAvvHWi6bqFxZXN1JFc2728cqm2lwnnsyxMW242MyMu/O0FSCQav6Prtlr8M0tjK0iwytBKrxtG8bjBKsrAEHBB5HIII4IoA0KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA84ufBXg/xd408QwyG5l1qN7C51CNomjG1DI1sVcoAw3CQh1YspGAy4ArrfCvhHT/B9ndW+nq4W6uXu5S7D5pGwCcAADoOgGTljlmYl9h4V03TfEGpa3BAy6nqMcUVzM0rtuWPdsUAnAA3t0A6/StegAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/9k=";
        private const string IMAGE_NOIMAGE = @"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";

        private Stream GetBinaryDataStream(string base64String)
        {
            if (base64String.Contains("http"))
            {
                Logger.Instance.Write("이미지 경로에 http가 포함됨");
                return new MemoryStream(Convert.FromBase64String(thumbnailPart1Data));
            }


            return new MemoryStream(Convert.FromBase64String(base64String));
        }


        /// <summary>
        /// 수평
        /// </summary>
        /// <param name="textHorizontal"></param>
        public Drawing.TextAlignmentTypeValues GetDrawingAlignment(TextAlignmentHorizontal textHorizontal)
        {
            Drawing.TextAlignmentTypeValues value = Drawing.TextAlignmentTypeValues.Left;
            switch (textHorizontal)
            {
                case TextAlignmentHorizontal.Left:
                    value = Drawing.TextAlignmentTypeValues.Left;
                    break;
                case TextAlignmentHorizontal.Center:
                    value = Drawing.TextAlignmentTypeValues.Center;
                    break;
                case TextAlignmentHorizontal.Right:
                    value = Drawing.TextAlignmentTypeValues.Right;
                    break;
            }
            return value;
        }

        /// <summary>
        /// 수직
        /// </summary>
        /// <param name="textVertical"></param>
        public Drawing.TextAnchoringTypeValues GetDrawingAnchoring(TextAlignmentVertical textVertical)
        {
            Drawing.TextAnchoringTypeValues value = Drawing.TextAnchoringTypeValues.Center;
            switch (textVertical)
            {
                case TextAlignmentVertical.Top:
                    value = Drawing.TextAnchoringTypeValues.Top;
                    break;
                case TextAlignmentVertical.Center:
                    value = Drawing.TextAnchoringTypeValues.Center;
                    break;
                case TextAlignmentVertical.Bottom:
                    value = Drawing.TextAnchoringTypeValues.Bottom;
                    break;
            }
            return value;
        }

        public Drawing.TextVerticalValues GetDrawingTextVertical(Model.Component.HomeTab.TextDirection direction)
        {
            Drawing.TextVerticalValues textVerticalValue = Drawing.TextVerticalValues.Horizontal;
            switch (direction)
            {
                case Model.Component.HomeTab.TextDirection.Vertical:
                    textVerticalValue = Drawing.TextVerticalValues.EastAsianVetical;
                    break;
                case Model.Component.HomeTab.TextDirection.RotateAllText90:
                    textVerticalValue = Drawing.TextVerticalValues.Vertical;
                    break;
                case Model.Component.HomeTab.TextDirection.RotateAllText270:
                    textVerticalValue = Drawing.TextVerticalValues.Vertical270;
                    break;
                case Model.Component.HomeTab.TextDirection.Stacked:
                    textVerticalValue = Drawing.TextVerticalValues.WordArtVertical;
                    break;
            }

            return textVerticalValue;
        }

        /// <summary>
        /// 수직
        /// </summary>
        /// <param name="textVertical"></param>
        public Wordprocessing.TableVerticalAlignmentValues GetWordprocessingTableVerticalAlignment(TextAlignmentVertical textVertical)
        {
            var value = Wordprocessing.TableVerticalAlignmentValues.Center;
            switch (textVertical)
            {
                case TextAlignmentVertical.Top:
                    value = Wordprocessing.TableVerticalAlignmentValues.Top;
                    break;
                case TextAlignmentVertical.Center:
                    value = Wordprocessing.TableVerticalAlignmentValues.Center;
                    break;
                case TextAlignmentVertical.Bottom:
                    value = Wordprocessing.TableVerticalAlignmentValues.Bottom;
                    break;
            }
            return value;
        }

        /// <summary>
        /// 수평
        /// </summary>
        /// <param name="textHorizontal"></param>
        public Wordprocessing.JustificationValues GetWordprocessingJustification(TextAlignmentHorizontal textHorizontal)
        {
            var value = Wordprocessing.JustificationValues.Left;
            switch (textHorizontal)
            {
                case TextAlignmentHorizontal.Left:
                    value = Wordprocessing.JustificationValues.Left;
                    break;
                case TextAlignmentHorizontal.Center:
                    value = Wordprocessing.JustificationValues.Center;
                    break;
                case TextAlignmentHorizontal.Right:
                    value = Wordprocessing.JustificationValues.Right;
                    break;
            }
            return value;
        }

        public Wordprocessing.TextDirectionValues GetWordpressingTextDirection(Model.Component.HomeTab.TextDirection direction)
        {
            var value = Wordprocessing.TextDirectionValues.LefToRightTopToBottom;
            switch (direction)
            {
                case Model.Component.HomeTab.TextDirection.Vertical:
                    // textVerticalValue = TextVerticalValues.EastAsianVetical;
                    break;
                case Model.Component.HomeTab.TextDirection.RotateAllText90:
                    //textVerticalValue = TextVerticalValues.Vertical;
                    break;
                case Model.Component.HomeTab.TextDirection.RotateAllText270:
                    // textVerticalValue = TextVerticalValues.Vertical270;
                    break;
                case Model.Component.HomeTab.TextDirection.Stacked:
                    value = Wordprocessing.TextDirectionValues.TopToBottomRightToLeftRotated;
                    break;
            }
            return value;
        }

        public Drawing.Transform2D GetDrawingTransfrom2D(int x, int y, int width, int height)
        {
            Drawing.Transform2D transform1 = new Drawing.Transform2D();
            Drawing.Offset offset1 = new Drawing.Offset() { X = x * EMUPPI, Y = y * EMUPPI };
            Drawing.Extents extents1 = new Drawing.Extents() { Cx = width * EMUPPI, Cy = height * EMUPPI };

            transform1.Append(offset1);
            transform1.Append(extents1);
            return transform1;
        }

        public Drawing.Outline GetDrawingOutline(float weight, string color)
        {
            var outline1 = new Drawing.Outline()
            {
                Width = (int)weight * EMUPPI,
            };
            Drawing.SolidFill solidFill2 = Common.Instance.GenerateSolidFill(color);
            outline1.Append(solidFill2);
            return outline1;
        }

        public Drawing.RunProperties GetDrawingRunProperty(OfficeFont font)
        {
            Drawing.TextUnderlineValues underlineValue = Drawing.TextUnderlineValues.None;
            if (font.UnderLine)
                underlineValue = Drawing.TextUnderlineValues.Single;

            Drawing.TextStrikeValues strikeValues = Drawing.TextStrikeValues.NoStrike;
            if (font.Strike)
                strikeValues = Drawing.TextStrikeValues.SingleStrike;

            var runProperties1 = new Drawing.RunProperties()
            {
                Language = "en-US",
                AlternativeLanguage = "ko-KR",
                Dirty = false,
                FontSize = (int)font.Size * 100,
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = underlineValue,
                Strike = strikeValues,
                Spacing = (int)(font.CharacterSpacing * 100)
            };

            Drawing.SolidFill solidFill1 = Common.Instance.GenerateSolidFill(font.Color);
            Drawing.LatinFont latinFont1 = new Drawing.LatinFont() { Typeface = font.Name };
            Drawing.EastAsianFont eastAsianFont1 = new Drawing.EastAsianFont() { Typeface = font.Name };


            runProperties1.Append(solidFill1);
            runProperties1.Append(latinFont1);
            runProperties1.Append(eastAsianFont1);

            return runProperties1;
        }

        public Wordprocessing.RunProperties GetWordRunProperty(OfficeFont font)
        {
            var runProperties = new Wordprocessing.RunProperties();

            #region 폰트명
            Wordprocessing.RunFonts runFonts4 = new Wordprocessing.RunFonts() 
            { 
                Hint = Wordprocessing.FontTypeHintValues.EastAsia, 
                Ascii = font.Name, 
                HighAnsi = font.Name, 
                EastAsia = font.Name 
            };
            runProperties.Append(runFonts4);
            #endregion

            #region 폰트사이즈
            var fontSize4 = new Wordprocessing.FontSize() { Val = (font.Size * 2).ToString() };
            runProperties.Append(fontSize4);
            #endregion

            #region 폰트옵션
            if (font.UnderLine)
            {
                var style = new Wordprocessing.Underline() { Val = Wordprocessing.UnderlineValues.Single };
                runProperties.Append(style);
            }
            if (font.Strike)
            {
                var style = new Wordprocessing.Strike();
                runProperties.Append(style);
            }
            if (font.Bold)
            {
                var style = new Wordprocessing.Bold();
                runProperties.Append(style);
            }
            if (font.Italic)
            {
                var style = new Wordprocessing.Italic();
                runProperties.Append(style);
            }
            #endregion

            #region 폰트컬러
            runProperties.Append(new Wordprocessing.Color() { Val = font.Color });
            #endregion

            #region 문자 간격
            var spacing4 = new Wordprocessing.Spacing() { Val = (int)(font.CharacterSpacing * 20) };
            runProperties.Append(spacing4);
            #endregion

            return runProperties;
        }
    }
}
