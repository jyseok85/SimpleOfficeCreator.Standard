﻿using DocumentFormat.OpenXml.Drawing;
using SimpleOfficeCreator.Stardard.Modules.Model.Component;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.PictureFormatTab;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.ShapeFormat;
using SimpleOfficeCreator.Stardard.Modules.Model.Component.TableDesignTab;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules.Model
{
    public enum Type
    { 
        None,
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
        public float Top { get; set; } = 0f;
        public float Right { get; set; } = 3.78f;
        public float Bottom { get; set; } = 0f;
    }
}