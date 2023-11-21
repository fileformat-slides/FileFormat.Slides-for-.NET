using System;
using System.Collections.Generic;
using System.Linq;
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

namespace FileFormat.Slides.Facade
{
    public class TextShapeFacade
    {
        private string _Text;
        private int _fontSize;
        private TextAlignment _alignment;
        private long _x;
        private long _y;
        private long _width;
        private long _height;
        private P.Shape _textBoxShape;
        private SlidePart _AssociatedSlidePart;// Store the P.Shape as a private field
        private int _ShapeIndex;
        private String _FontFamily = "Calibri";
        private String _TextColor = "#000000";
        public string Text { get => _Text; set => _Text = value; }
        // Public property to set and get font size in points
        public int FontSize
        {
            get
            {
                // Convert fontSizeTwentieths to points
                return (int)_fontSize / 20 / 2;
            }
            set
            {
                // Convert points to twentieths of a point and store in fontSizeTwentieths
                _fontSize = (int)(value * 20 * 2);
            }
        }
        public TextAlignment Alignment { get => _alignment; set => _alignment = value; }
        public long X { get => _x; set => _x = value; }
        public long Y { get => _y; set => _y = value; }
        public long Width { get => _width; set => _width = value; }
        public long Height { get => _height; set => _height = value; }
        public P.Shape TextBoxShape { get => _textBoxShape; set => _textBoxShape = value; }
        protected  SlidePart AssociatedSlidePart { get => _AssociatedSlidePart; set => _AssociatedSlidePart = value; }
        public  int ShapeIndex { get => _ShapeIndex; set => _ShapeIndex = value; }
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        public string TextColor { get => _TextColor; set => _TextColor = value; }

        public TextShapeFacade ()
        {
            // Set default values or initialization logic if needed
            FontSize = 12; // Example default font size
            Alignment = TextAlignment.Left; // Example default alignment
        }

        public TextShapeFacade WithText (string text)
        {
            Text = text;
            return this;
        }

        public TextShapeFacade WithFontSize (int fontSize)
        {
            FontSize = fontSize;
            return this;
        }
        public TextShapeFacade WithFontFamily (String fontfamily)
        {
            FontFamily = fontfamily;
            return this;
        }
        public TextShapeFacade WithTextColor (String textColor)
        {
            TextColor = textColor;
            return this;
        }
        public TextShapeFacade WithAlignment (TextAlignment alignment)
        {
            Alignment = alignment;
            return this;
        }

        public TextShapeFacade WithPosition (long x, long y)
        {
            X = x;
            Y = y;
            return this;
        }

        public TextShapeFacade WithSize (long width, long height)
        {
            Width = width;
            Height = height;
            return this;
        }

        public P.Shape CreateShape ()
        {
            TextAlignmentTypeValues alignmentType = ConvertAlignmentToTypeValues(Alignment);

            TextBoxShape = new P.Shape(
               new P.NonVisualShapeProperties(
                   new P.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Box 1" },
                   new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                   new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                new P.ShapeProperties(new D.Transform2D()
                {
                    Offset = new D.Offset() { X = X, Y = Y },
                    Extents = new D.Extents() { Cx = Width, Cy = Height }
                }),
               new P.TextBody(
                   new BodyProperties(),
                   new ListStyle(),
                   new Paragraph(
                       new ParagraphProperties() { Alignment = alignmentType },
                       new Run(
                           new RunProperties(new SolidFill(new RgbColorModelHex() { Val = _TextColor }),
                           new LatinFont() { Typeface= _FontFamily }) { FontSize = _fontSize, Dirty = false },
                             new Text() { Text = Text } 
                       )
                   )
               )
           );

            return TextBoxShape;
        }


        public void UpdateShape ()
        {
            if (TextBoxShape == null)
            {
                throw new InvalidOperationException("Shape has not been created yet. Call CreateShape() first.");
            }

            TextAlignmentTypeValues alignmentType = ConvertAlignmentToTypeValues(Alignment);

            // Update the properties of the existing shape
            TextBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = (UInt32Value)5U;
            TextBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name = "Text Box 1";
            TextBoxShape.NonVisualShapeProperties.NonVisualShapeDrawingProperties = new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true });
            TextBoxShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties(new PlaceholderShape());

            TextBoxShape.ShapeProperties.Transform2D = new D.Transform2D()
            {
                Offset = new D.Offset() { X = X, Y = Y },
                Extents = new D.Extents() { Cx = Width, Cy = Height }
            };
            
            var existingParagraphText = TextBoxShape.TextBody.Descendants<Run>().FirstOrDefault();
            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().RemoveAllChildren();            
            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(new ParagraphProperties() { Alignment = alignmentType });
            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(existingParagraphText);

            var runProperties = TextBoxShape.TextBody.Descendants<RunProperties>().FirstOrDefault();

            if (_fontSize != 0)
            {
                runProperties.FontSize = _fontSize;
            }
            else if(runProperties.FontSize != null)
            {

            }
            else
            {
                // Extract font size from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];

                if (defaultRunProperties != null)
                {
                    runProperties.FontSize = defaultRunProperties.FontSize;
                }
            }
            var latinFont = runProperties.Elements<LatinFont>().FirstOrDefault();

            if (!string.IsNullOrEmpty(_FontFamily))
            {
                latinFont = new LatinFont() { Typeface = _FontFamily };
            }
            else if (latinFont != null)
            {
                
            }
            else
            {
                // Extract font family from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];
                var defaultLatinFont= defaultRunProperties.Elements<LatinFont>().FirstOrDefault();

                if (defaultRunProperties != null && defaultLatinFont != null)
                {
                    latinFont = new LatinFont() { Typeface = defaultLatinFont.Typeface };
                }
            }
          
            var solidFill = runProperties.Elements<SolidFill>().FirstOrDefault();

            if (!string.IsNullOrEmpty(_TextColor))
            {
                // Clear existing font-related properties
                runProperties.Elements<LatinFont>().ToList().ForEach(l => l.Remove());

                // Set text color using hex code (e.g., #FF0000 for red)
                solidFill = new SolidFill(new RgbColorModelHex() { Val = _TextColor });
                runProperties.AppendChild(solidFill);
            }
            else if (solidFill != null)
            {
               
            }
            else
            {
                // Extract color from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];

                if (defaultRunProperties != null)
                {
                    // Clear existing font-related properties
                    runProperties.Elements<LatinFont>().ToList().ForEach(l => l.Remove());

                    // Set text color using hex code (e.g., #FF0000 for red)
                    solidFill = new SolidFill(new RgbColorModelHex() { Val = _TextColor });
                    runProperties.AppendChild(solidFill);
                }
            }


            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Elements<Run>().FirstOrDefault().Text.Text = Text;
        }
       

        public void RemoveShape (SlidePart slidePart)
        {
            // Ensure slidePart is not null
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart cannot be null.");
            }

            // Find the ShapeTree in CommonSlideData
            CommonSlideData commonSlideData = slidePart.Slide.CommonSlideData;
            if (commonSlideData != null && commonSlideData.ShapeTree != null)
            {
                // Remove the specified shape from the ShapeTree
                var shapesToRemove = commonSlideData.ShapeTree.Elements<P.Shape>().Where(shape => shape == TextBoxShape).ToList();

                foreach (var shape in shapesToRemove)
                {
                    shape.Remove();
                }
            }
        }
        public void RemoveShape (P.Shape shape)
        {
            shape.Remove();
        }

        // Method to populate List<TextShapeFacade> from a collection of P.Shape
        public static List<TextShapeFacade> PopulateTextShapes (SlidePart slidePart)
        {
            IEnumerable<P.Shape> shapes = slidePart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>();
            var textShapes = new List<TextShapeFacade>();
            var shapeIndex = 0;
            foreach (var shape in shapes)
            {

                var textShapeFacade = new TextShapeFacade
                {
                    TextBoxShape = shape, // Store the P.Shape in the private field
                    Text = GetTextFromTextShape(shape),
                    FontSize = GetFontSizeFromTextShape(shape),
                    TextColor = GetColorFromTextShape(shape),
                    Alignment = GetAlignmentFromTextShape(shape),
                    FontFamily = GetFontFamilyFromTextShape(shape),
                    X = GetXFromShape(shape),
                    Y = GetYFromShape(shape),
                    Width = GetWidthFromShape(shape),
                    Height = GetHeightFromShape(shape),
                    AssociatedSlidePart = slidePart,
                    ShapeIndex=shapeIndex
                };

                textShapes.Add(textShapeFacade);
                shapeIndex += 1;
            }

            return textShapes;
        }

        private static string GetTextFromTextShape (P.Shape textShape)
        {
            if (textShape.TextBody != null)
            {
                return textShape.TextBody.Descendants<Text>().FirstOrDefault()?.Text;
            }
            return null;
        }

        private static int GetFontSizeFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    return defaultRunProperties.FontSize;
                }
            }

            return 0;
        }
        private static string GetFontFamilyFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    var latinFont = defaultRunProperties.Elements<LatinFont>().FirstOrDefault();

                    if (latinFont != null)
                    {
                        return latinFont.Typeface;
                    }
                }
            }

            return null; // or an appropriate default value for FontFamily
        }
        private static string GetColorFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    var solidFill = defaultRunProperties.Elements<SolidFill>().FirstOrDefault();

                    if (solidFill != null)
                    {
                        var rgbColor = solidFill.Elements<RgbColorModelHex>().FirstOrDefault();

                        if (rgbColor != null)
                        {
                            return rgbColor.Val;
                        }
                    }
                }
            }

            return null; // or an appropriate default value for color code
        }

        private static TextAlignment GetAlignmentFromTextShape (P.Shape textShape)
        {
            var alignment = textShape.TextBody?.Descendants<Paragraph>().FirstOrDefault();
            if (alignment !=null)
            {
                alignment = null;
            }
            TextAlignmentTypeValues alignmentType = textShape.TextBody?.Descendants<ParagraphProperties>().FirstOrDefault()?.Alignment ?? TextAlignmentTypeValues.Left;
            return ConvertAlignmentFromTypeValues(alignmentType);
        }

        private static long GetXFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.X ?? 0;
        }

        private static long GetYFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.Y ?? 0;
        }

        private static long GetWidthFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0;
        }

        private static long GetHeightFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0;
        }

        private static TextAlignmentTypeValues ConvertAlignmentToTypeValues (TextAlignment alignment)
        {
            switch (alignment)
            {
                case TextAlignment.Left:
                    return TextAlignmentTypeValues.Left;
                case TextAlignment.Center:
                    return TextAlignmentTypeValues.Center;
                case TextAlignment.Right:
                    return TextAlignmentTypeValues.Right;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
            }
        }

        private static TextAlignment ConvertAlignmentFromTypeValues (TextAlignmentTypeValues alignmentType)
        {
            switch (alignmentType)
            {
                case TextAlignmentTypeValues.Left:
                    return TextAlignment.Left;
                case TextAlignmentTypeValues.Center:
                    return TextAlignment.Center;
                case TextAlignmentTypeValues.Right:
                    return TextAlignment.Right;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignmentType), alignmentType, null);
            }
        }
    }
}
