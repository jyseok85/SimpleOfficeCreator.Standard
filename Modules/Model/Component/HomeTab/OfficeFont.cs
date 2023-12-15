namespace SimpleOfficeCreator.Standard.Modules.Model.Component.HomeTab
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
        /// 문자 간격
        /// UI상에서 표현되는 좁게 옵션은 - 값으로 입력하면 됩니다.
        /// </summary>
        public float CharacterSpacing { get; set; } = 0;


        /// <summary>
        /// HexColorValue
        /// </summary>
        public string Color { get; set; } = "000000";

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
