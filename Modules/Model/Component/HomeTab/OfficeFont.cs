using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab
{
    public enum Spacing
    {
        VeryTight = -300,
        Tight = -150,
        Normal = 0,
        Loose = 300,
        VeryLoose = 600,
        Custom
    }

    public class OfficeFont
    {
        //구현안된 항목
        //1.그림자,
        //2.대소문자 바꾸기

        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// pt
        /// </summary>
        public float Size { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public bool UnderLine { get; set; }

        public bool Strike { get; set; }

        /// <summary>
        /// 자간
        /// </summary>
        public float SpacingValue { get; set; } = 0;

        /// <summary>
        /// HexColorValue
        /// </summary>
        public string Color { get; set; } = "000000";



        /// <summary>
        /// 자간을 설정합니다. 
        /// </summary>
        /// <param name="spacing">Office UI 에 있는 기본 DropDown 값</param>
        /// <param name="point">Custom 일때 입력할 값\n(pt * 100) 수치를 입력</param>
        public void SetSpacing(Spacing spacing, int point = 0)
        {
            if (spacing == Spacing.Custom)
            {
                SpacingValue = point;
            }
            else
                SpacingValue = (int)spacing;
        }

        /// <summary>
        /// 오피스에서는 RGB Hex 값이 사용되며, 투명은 NoFill 속성으로 처리된다.
        /// </summary>
        /// <param name="text">컬러 문자열</param
        public void SetColor(string text)
        {
            if (text.Contains("#"))
            {
                Color = text;
            }
            else if (text.ToLower().Equals("transparent"))
            {
                Color = "transparent";
            }
            else
            {
                System.Drawing.Color myColor = System.Drawing.Color.FromName(text);
                string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
                Color = hex;
            }
        }
    }
}
