using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    public class Common
    {
        private Common() { }
        //private static 인스턴스 객체
        private static readonly Lazy<Common> _instance = new Lazy<Common>(() => new Common());
        //public static 의 객체반환 함수
        public static Common Instance { get { return _instance.Value; } }


        /// <summary>
        /// 배경색 컨트롤을 생성한다.
        /// </summary>
        /// <returns></returns>
        public SolidFill GenerateSolidFill(string color)
        {
            var borderColor = color;
            //투명으로 들어왔다면 그냥 흰색으로 변경한다. 대신 추후 Alpha 컴포넌트를 추가한다. 
            if (borderColor == "transparent")
                borderColor = "FFFFFF";

            SolidFill solidFill = new SolidFill();
            RgbColorModelHex rgbBackColor = new RgbColorModelHex() { Val = borderColor };
            if (color == "transparent")
            {
                Alpha alpha = new Alpha() { Val = 0 };
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
            if(text == null)
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
            else if(text.Contains(","))
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
    }   
}
