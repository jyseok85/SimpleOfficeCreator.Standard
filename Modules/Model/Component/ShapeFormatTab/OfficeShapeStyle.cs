namespace SimpleOfficeCreator.Standard.Modules.Model.Component.ShapeFormat
{
    public class OfficeShapeStyle
    {
        public bool UseFill { get; set; } = false;
        /// <summary>
        /// 도형 채우기 색
        /// </summary>
        public string FillColor { get; set; } = "trasnparent";

        /// <summary>
        /// 도형 윤곽선
        /// </summary>
        public bool UseOutline { get; set; } = false;
        public float OutlineWeight { get; set; } = 1;
        public string OutlineDashes { get; set; } = "Solid";
        public string OutlineColor { get; set; } = "000000";

        public string ShapeTypeValue { get; set; } = "rectangle";
    }
}
