using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.EMMA;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleOfficeCreator.Stardard.Modules.GeneratedCode
{
    internal class testclass
    {
        public Run TEST()
        {
            Run run104 = new Run() { RsidRunProperties = "00D5165D" };

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "돋움", HighAnsi = "돋움", EastAsia = "돋움" };
            NoProof noProof4 = new NoProof();
            Color color100 = new Color() { Val = "000000" };
            Spacing spacing94 = new Spacing() { Val = -16 };
            FontSize fontSize15 = new FontSize() { Val = "16" };

            runProperties100.Append(runFonts102);
            runProperties100.Append(noProof4);
            runProperties100.Append(color100);
            runProperties100.Append(spacing94);
            runProperties100.Append(fontSize15);

            AlternateContent alternateContent3 = new AlternateContent();

            AlternateContentChoice alternateContentChoice3 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing4 = new Drawing();

            Anchor anchor4 = new Anchor() { DistanceFromTop = (UInt32Value)45720U, DistanceFromBottom = (UInt32Value)45720U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251664384U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "6FC46515", AnchorId = "15C98A1E" };
            SimplePosition simplePosition4 = new SimplePosition() { X = 0L, Y = 0L };

            HorizontalPosition horizontalPosition4 = new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.Column };
            PositionOffset positionOffset7 = new PositionOffset();
            positionOffset7.Text = "6096000";

            horizontalPosition4.Append(positionOffset7);

            VerticalPosition verticalPosition4 = new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Paragraph };
            PositionOffset positionOffset8 = new PositionOffset();
            positionOffset8.Text = "10191750";

            verticalPosition4.Append(positionOffset8);
            Extent extent4 = new Extent() { Cx = 1066800L, Cy = 304800L };
            EffectExtent effectExtent4 = new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            WrapNone wrapNone4 = new WrapNone();
            DocProperties docProperties4 = new DocProperties() { Id = (UInt32Value)2U, Name = "텍스트 상자 2" };

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties4 = new NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks4 = new A.GraphicFrameLocks();
            graphicFrameLocks4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties4.Append(graphicFrameLocks4);

            A.Graphic graphic4 = new A.Graphic();
            graphic4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData4 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties3.Append(shapeLocks3);

            Wps.ShapeProperties shapeProperties4 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents4 = new A.Extents() { Cx = 1066800L, Cy = 304800L };

            transform2D4.Append(offset4);
            transform2D4.Append(extents4);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill3.Append(rgbColorModelHex3);

            A.Outline outline3 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill4.Append(rgbColorModelHex4);
            A.Miter miter3 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd3 = new A.HeadEnd();
            A.TailEnd tailEnd3 = new A.TailEnd();

            outline3.Append(solidFill4);
            outline3.Append(miter3);
            outline3.Append(headEnd3);
            outline3.Append(tailEnd3);

            shapeProperties4.Append(transform2D4);
            shapeProperties4.Append(presetGeometry4);
            shapeProperties4.Append(solidFill3);
            shapeProperties4.Append(outline3);

            Wps.TextBoxInfo2 textBoxInfo23 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent5 = new TextBoxContent();

            Paragraph paragraph141 = new Paragraph() { RsidParagraphAddition = "00D5165D", RsidParagraphProperties = "00D5165D", RsidRunAdditionDefault = "00D5165D" };

            ParagraphProperties paragraphProperties104 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts103 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties5.Append(runFonts103);

            paragraphProperties104.Append(paragraphMarkRunProperties5);

            Run run105 = new Run();
            Text text87 = new Text();
            text87.Text = "2";

            run105.Append(text87);

            Run run106 = new Run();

            RunProperties runProperties101 = new RunProperties();
            RunFonts runFonts104 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            runProperties101.Append(runFonts104);
            Text text88 = new Text();
            text88.Text = "번 텍";

            run106.Append(runProperties101);
            run106.Append(text88);

            Run run107 = new Run();
            Text text89 = new Text();
            text89.Text = "스트";

            run107.Append(text89);

            paragraph141.Append(paragraphProperties104);
            paragraph141.Append(run105);
            paragraph141.Append(run106);
            paragraph141.Append(run107);

            textBoxContent5.Append(paragraph141);

            textBoxInfo23.Append(textBoxContent5);

            Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false };
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            textBodyProperties3.Append(noAutoFit3);

            wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
            wordprocessingShape3.Append(shapeProperties4);
            wordprocessingShape3.Append(textBoxInfo23);
            wordprocessingShape3.Append(textBodyProperties3);

            graphicData4.Append(wordprocessingShape3);

            graphic4.Append(graphicData4);

            Wp14.RelativeWidth relativeWidth4 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth4 = new Wp14.PercentageWidth();
            percentageWidth4.Text = "0";

            relativeWidth4.Append(percentageWidth4);

            Wp14.RelativeHeight relativeHeight4 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight4 = new Wp14.PercentageHeight();
            percentageHeight4.Text = "0";

            relativeHeight4.Append(percentageHeight4);

            anchor4.Append(simplePosition4);
            anchor4.Append(horizontalPosition4);
            anchor4.Append(verticalPosition4);
            anchor4.Append(extent4);
            anchor4.Append(effectExtent4);
            anchor4.Append(wrapNone4);
            anchor4.Append(docProperties4);
            anchor4.Append(nonVisualGraphicFrameDrawingProperties4);
            anchor4.Append(graphic4);
            anchor4.Append(relativeWidth4);
            anchor4.Append(relativeHeight4);

            drawing4.Append(anchor4);

            alternateContentChoice3.Append(drawing4);

            AlternateContentFallback alternateContentFallback3 = new AlternateContentFallback();

            Picture picture4 = new Picture();

            V.Shape shape3 = new V.Shape() { Id = "_x0000_s1028", Style = "position:absolute;left:0;text-align:left;margin-left:480pt;margin-top:802.5pt;width:84pt;height:24pt;z-index:251664384;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", Type = "#_x0000_t202", EncodedPackage = "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\n" };
            shape3.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "15C98A1E"));

            V.TextBox textBox3 = new V.TextBox();

            TextBoxContent textBoxContent6 = new TextBoxContent();

            Paragraph paragraph142 = new Paragraph() { RsidParagraphAddition = "00D5165D", RsidParagraphProperties = "00D5165D", RsidRunAdditionDefault = "00D5165D" };

            ParagraphProperties paragraphProperties105 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts105 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties6.Append(runFonts105);

            paragraphProperties105.Append(paragraphMarkRunProperties6);

            Run run108 = new Run();
            Text text90 = new Text();
            text90.Text = "2";

            run108.Append(text90);

            Run run109 = new Run();

            RunProperties runProperties102 = new RunProperties();
            RunFonts runFonts106 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            runProperties102.Append(runFonts106);
            Text text91 = new Text();
            text91.Text = "번 텍";

            run109.Append(runProperties102);
            run109.Append(text91);

            Run run110 = new Run();
            Text text92 = new Text();
            text92.Text = "스트";

            run110.Append(text92);

            paragraph142.Append(paragraphProperties105);
            paragraph142.Append(run108);
            paragraph142.Append(run109);
            paragraph142.Append(run110);

            textBoxContent6.Append(paragraph142);

            textBox3.Append(textBoxContent6);

            shape3.Append(textBox3);

            picture4.Append(shape3);

            alternateContentFallback3.Append(picture4);
            alternateContent3.Append(alternateContentFallback3);

            alternateContent3.Append(alternateContentChoice3);

            run104.Append(runProperties100);
            run104.Append(alternateContent3);
            return run104;
        }


    }
}
