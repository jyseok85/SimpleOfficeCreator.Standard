namespace SimpleOfficeCreator.Stardard.Modules.Model.Component.TableDesignTab
{
    public class OfficeTableStyles
    {
        /// <summary>
        /// 배경색 사용 유무
        /// </summary>
        public bool UseShading { get; set; } = false;

        public string ShadingColor { get; set; } = "transparent";

        public Border Top { get; set; } = new Border();
        public Border Right { get; set; } = new Border();
        public Border Bottom { get; set; } = new Border();
        public Border Left { get; set; } = new Border();
    }
    public class Border
    {
        public bool Draw { get; set; } = false;
        public string Color { get; set; } = "Black";
        public float Weight { get; set; } = 1;
        public string Style { get; set; } = "Solid";
    }
}
