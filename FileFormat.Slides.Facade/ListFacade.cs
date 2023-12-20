using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using FileFormat.Slides.Common.Enumerations;
using FileFormat.Slides.Common;
using System.Collections.Generic;
using System;

namespace FileFormat.Slides.Facade
{
    
    public class ListFacade
    {
        private ListStyle _ListStyle;
        public ListFacade ()
        {
            _ListStyle = new ListStyle();
        }
        private Paragraph AddListItem (String Text, String _TextColor, String _FontFamily, int _fontSize)
        {
            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties() { LeftMargin = 342900, Indent = -342900, Alignment = TextAlignmentTypeValues.Left };
            BulletFont bulletFont1 = new BulletFont() { Typeface = _FontFamily, Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            CharacterBullet characterBullet1 = new CharacterBullet() { Char = "•" };

            paragraphProperties1.Append(bulletFont1);
            paragraphProperties1.Append(characterBullet1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties(new SolidFill(new RgbColorModelHex() { Val = _TextColor }),
                 new LatinFont() { Typeface = _FontFamily }) { FontSize = _fontSize, Language = "en-US", Dirty = false };
            Text text1 = new Text();
            text1.Text = Text;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            return paragraph1;
        }
        public P.TextBody CreateList (List<String> listItems,String _textColor, String _fontFamily, int _fontSize, P.TextBody body)
        {
            body.Append(new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center });
            body.Append(_ListStyle);
            foreach(var text in listItems)
            {
                body.Append(AddListItem(text, _textColor, _fontFamily, _fontSize));
            }
            return body;
        }

    }
}
