using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules.Model.Component.PictureFormatTab
{

    public class OfficePictureStyle
    {
        /// <summary>
        /// 일관성 없이 NoFill, NoOutline, Border, Outline 섞어서 사용하고 있는데.
        /// 다 Office UI의 텍스트랑 매칭한 결과다.. 거기서부터 섞어서 쓰고 있다.
        /// </summary>
        public bool NoOutline { get; set; } = true;
        public float Weight { get; set; } = 1;
        public string Dashes { get; set; } = "Solid";
        public string Color { get; set; } = "000000";
    }
}
