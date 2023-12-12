using System.ComponentModel;

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
    public enum TextAlignmentHorizontal
    {
        Left,
        Center,
        Right
    }
    public enum TextAlignmentVertical
    {
        Top,
        Center,
        Bottom
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
        public TextAlignmentHorizontal AlignmentHorizontal { get; set; } = TextAlignmentHorizontal.Left;
        //public TextAlignmentTypeValues AlignmentHorizontal { get; set; } = TextAlignmentTypeValues.Left;
        /// <summary>
        /// 수직정렬
        /// </summary>
        public TextAlignmentVertical AlignmentVertical { get; set; } = TextAlignmentVertical.Top;
        //public TextAnchoringTypeValues AlignmentVertical { get; set; } = TextAnchoringTypeValues.Top;
        /// <summary>
        /// 오피스 UI 와 동일하게 처리됨. 실제 Openxml emum과는 명칭이 다름.
        /// </summary>
        [DefaultValue(TextDirection.Horizontal)]
        public TextDirection TextDirection { get; set; } = TextDirection.Horizontal;

        /// <summary>
        /// 줄간격 (pt 단위 이며, 폰트사이즈의 1.2배가 PPT에서 기본값(1줄)과 동일하다. WORD의 경우 폰트에 따라서 값이 달라진다.
        /// Word의 경우 최소값이 0.7pt 이상이다. 
        /// 모델의 기본값(0) 일경우 옵션을 적용하지 않는다.
        /// </summary>
        public float LineSpacing { get; set; } = 0;

        /// <summary> 
        /// Word 전용
        /// </summary>
        public bool TableCellFitText { get; set; } = false;
    }
}
