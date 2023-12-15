using System.ComponentModel;

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
        /// Point
        /// </summary>
        public float Size { get; set; } = 10;
        
        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public bool UnderLine { get; set; }

        public bool Strike { get; set; }

        /// <summary>
        /// 문자 간격
        /// </summary>
        [DefaultValue(0)]
        public float CharacterSpacing { get; set; } = 0;


        /// <summary>
        /// HexColorValue
        /// </summary>
        public string Color { get; set; } = "000000";
    }
}
