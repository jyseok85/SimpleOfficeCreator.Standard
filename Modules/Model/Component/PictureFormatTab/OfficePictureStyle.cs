﻿namespace SimpleOfficeCreator.Standard.Modules.Model.Component.PictureFormatTab
{

    public class OfficePictureStyle
    {
        /// <summary>
        /// 일관성 없이 NoFill, NoOutline, Border, Outline 섞어서 사용하고 있는데.
        /// 다 Office UI의 텍스트랑 매칭한 결과다.. 거기서부터 섞어서 쓰고 있다.
        /// </summary>
        public bool NoOutline { get; set; } = false;
        public float Weight { get; set; } = 1;
        /// <summary>
        /// 사용가능 - SOLID,DOT,DASH,DASHDOT,DASHDOTDOT
        /// </summary>
        public string Dashes { get; set; } = "Solid";
        public string Color { get; set; } = "000000";
    }
}