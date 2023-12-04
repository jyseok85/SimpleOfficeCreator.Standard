using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOfficeCreator.Stardard.Modules.Model.Component.HomeTab
{
    public enum TextDirection
    {
        Horizontal,
        Vertical,
        RotateAllText90,
        RotateAllText270,
        Stacked
    }

    public class OfficeParagraph
    {

        //구현안된 항목
        //글머리 번호
        //번호 매기기
        //목록 수준 늘리기, 줄이기
        //열 추가 제거
        //SmartArt로 변환

        /// <summary>
        /// 수평정렬
        /// </summary>
        public TextAlignmentTypeValues AlignmentHorizontal { get; set; } = TextAlignmentTypeValues.Left;
        /// <summary>
        /// 수직정렬
        /// </summary>
        public TextAnchoringTypeValues AlignmentVertical { get; set; } = TextAnchoringTypeValues.Top;
        /// <summary>
        /// 오피스 UI 와 동일하게 처리됨. 실제 Openxml emum과는 명칭이 다름.
        /// </summary>
        [DefaultValue(TextDirection.Horizontal)]
        public TextDirection TextDirection { get; set; } = TextDirection.Horizontal;

        /// <summary>
        /// 줄간격
        /// </summary>
        public int LineSpacing { get; set; }



    }
}
