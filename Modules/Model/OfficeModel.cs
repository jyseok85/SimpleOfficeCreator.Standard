using SimpleOfficeCreator.Stardard.Modules.Model.Component;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.PictureFormatTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.ShapeFormat;

namespace SimpleOfficeCreator.Stardard.Modules.Model
{
    public enum Type
    {
        None,
        Paper,
        Table,
        TableCell,
        Picture,
        TextBox,
        Shape
    }

    public class OfficeModel
    {
        public OfficeModel(string name, string uid = "", string parentUid = "", bool visible = true)
        {
            this.Name = name;
            this.UID = uid;
            this.PARENT_UID = parentUid;
            this.Visible = visible;
        }
        public string PARENT_UID { get; set; } = string.Empty;
        public string UID { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;
        public Type Type { get; set; } = Type.None;

        public bool Visible { get; set; } = true;

        public string Text { get; set; } = string.Empty;


        public OfficeFont Font { get; set; } = new OfficeFont();
        public OfficeParagraph Paragraph { get; set; } = new OfficeParagraph();
        /// <summary>
        /// Pixel 단위
        /// </summary>
        public Margin Margin { get; set; } = new Margin();

        public Rectangle Rect = new Rectangle();

        public OfficeShapeStyle ShapeStyle { get; set; } = null;

        public OfficeTableInfo TableInfo { get; set; } = null;

        public OfficePictureStyle PictureStyle { get; set; } = null;

        public PaperInfo PaperInfo { get; set; } = null;
    }


    /// <summary>
    /// 마진을 포함한 절대값으로 설정합니다. 
    /// </summary>
    public class Rectangle
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        /// <summary>
        /// 필수값 아님
        /// </summary>
        public int Right { get; set; }
        /// <summary>
        /// 필수값 아님
        /// </summary>
        public int Bottom { get; set; }
    }

    public class Margin
    {
        public float Left { get; set; } = 3.78f;
        public float Right { get; set; } = 3.78f;
        //실제 DR은 0.1cm가 할당되어 있지만, 문서 변환시 0으로 해야 DR과 비슷하게 나온다. 
        public float Top { get; set; } = 3.78f;
        public float Bottom { get; set; } = 3.78f;
    }

    public class PaperInfo
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsLandscape { get; set; }
    }

}
