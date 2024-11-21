using System;
using System.Collections.Generic;
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
namespace FileFormat.Slides.Facade
{
    public class SlideFacade
    {
        
        private Slide _PresentationSlide;
        private SlidePart _SlidePart; 

        private String _RelationshipId;
        private int _SlideIndex;
        private String _BackgroundColor;

        private List<TextShapeFacade> _TextShapeFacades = null;

        private List<RectangleShapeFacade> _RectangleShapeFacades = null;

        private List<TriangleShapeFacade> _TriangleShapeFacades = null;

        private List<DiamondShapeFacade> _DiamondShapeFacades = null;

        private List<LineFacade> _LineFacades = null;

        private List<CurvedLineFacade> _CurvedLineFacades = null;

        private List<ArrowFacade> _ArrowFacades = null;

        private List<DoubleArrowFacade> _DoubleArrowFacades = null;

        private List<DoubleBraceFacade> _DoubleBraceFacades = null;

        private List<DoubleBracketFacade> _DoubleBracketFacades = null;

        private List<PentagonFacade> _PentagonFacades = null;

        private List<HexagonFacade> _HexagonFacades = null;

        private List<TrapezoidFacade> _TrapezoidFacades = null;

        private List<PieFacade> _PieFacades = null;

        private List<CircleShapeFacade> _CircleShapeFacades = null;

        private List<ImageFacade> _ImagesFacade = null;

        private List<TableFacade> _TableFacades = null;

        private CommentAuthorsPart _CommentAuthorPart;

        private SlideCommentsPart _CommentPart;

        private NotesSlidePart _NotesPart;

        public Slide PresentationSlide { get => _PresentationSlide; set => _PresentationSlide = value; }
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }
        public SlidePart SlidePart { get => _SlidePart; set => _SlidePart = value; }
        public List<TextShapeFacade> TextShapeFacades { get => _TextShapeFacades; set => _TextShapeFacades = value; }
        public int SlideIndex { get => _SlideIndex; set => _SlideIndex = value; }
        public List<ImageFacade> ImagesFacade { get => _ImagesFacade; set => _ImagesFacade = value; }
        public String BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        public List<TableFacade> TableFacades { get => _TableFacades; set => _TableFacades = value; }
        public CommentAuthorsPart CommentAuthorPart { get => _CommentAuthorPart; set => _CommentAuthorPart = value; }
        public SlideCommentsPart CommentPart { get => _CommentPart; set => _CommentPart = value; }
        public NotesSlidePart NotesPart { get => _NotesPart; set => _NotesPart = value; }
        public List<RectangleShapeFacade> RectangleShapeFacades { get => _RectangleShapeFacades; set => _RectangleShapeFacades = value; }
        public List<CircleShapeFacade> CircleShapeFacades { get => _CircleShapeFacades; set => _CircleShapeFacades = value; }
        public List<TriangleShapeFacade> TriangleShapeFacades { get => _TriangleShapeFacades; set => _TriangleShapeFacades = value; }
        public List<DiamondShapeFacade> DiamondShapeFacades { get => _DiamondShapeFacades; set => _DiamondShapeFacades = value; }
        public List<LineFacade> LineFacades { get => _LineFacades; set => _LineFacades = value; }
        public List<ArrowFacade> ArrowFacades { get => _ArrowFacades; set => _ArrowFacades = value; }
        public List<DoubleArrowFacade> DoubleArrowFacades { get => _DoubleArrowFacades; set => _DoubleArrowFacades = value; }
        public List<CurvedLineFacade> CurvedLineFacades { get => _CurvedLineFacades; set => _CurvedLineFacades = value; }
        public List<DoubleBraceFacade> DoubleBraceFacades { get => _DoubleBraceFacades; set => _DoubleBraceFacades = value; }
        public List<PentagonFacade> PentagonFacades { get => _PentagonFacades; set => _PentagonFacades = value; }
        public List<DoubleBracketFacade> DoubleBracketFacades { get => _DoubleBracketFacades; set => _DoubleBracketFacades = value; }
        public List<HexagonFacade> HexagonFacades { get => _HexagonFacades; set => _HexagonFacades = value; }
        public List<TrapezoidFacade> TrapezoidFacades { get => _TrapezoidFacades; set => _TrapezoidFacades = value; }
        public List<PieFacade> PieFacades { get => _PieFacades; set => _PieFacades = value; }

        public SlideFacade (bool isNewSlide)
        {
            if (isNewSlide)
            {
                Utility.NextIndex += 1;
                _RelationshipId = Utility.GetUniqueRelationshipId();
                var _PresentationPart = PresentationDocumentFacade.getInstance().GetPresentationPart();
                var _SlideIdList = _PresentationPart.Presentation.SlideIdList;
                _SlideIdList.Append(new SlideId() { Id = (UInt32Value)Utility.GetRandomSlideId(), RelationshipId = _RelationshipId });
               
                _PresentationSlide = new Slide(
                    new CommonSlideData(
                         new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()))),
                    new ColorMapOverride(new MasterColorMapping()));
                _SlidePart = _PresentationPart.AddNewPart<SlidePart>(_RelationshipId);

                if (PresentationDocumentFacade.getInstance().PresentationSlideLayoutParts.Count != 0)
                    _SlidePart.AddPart(PresentationDocumentFacade.getInstance().PresentationSlideLayoutParts[0]);
            }
         }
        public void SetSlideBackground (string color)
        {
            // Check if color is not null or empty
            if (!string.IsNullOrEmpty(color))
            {
                // Check if there is already a Background element
                Background existingBackground = _PresentationSlide.CommonSlideData.Elements<Background>().FirstOrDefault();

                // If an existing background is found, remove it
                if (existingBackground != null)
                {
                    _PresentationSlide.CommonSlideData.RemoveChild(existingBackground);
                }

                // Create a new Background element with the specified color
                Background newBackground = new Background();
                BackgroundProperties backgroundProperties = new BackgroundProperties();
                SolidFill solidFill = new SolidFill();
                RgbColorModelHex rgbColorModelHex = new RgbColorModelHex() { Val = color };
                solidFill.Append(rgbColorModelHex);
                backgroundProperties.Append(solidFill);
                newBackground.Append(backgroundProperties);

                // Insert the new Background element before the ShapeTree
                _PresentationSlide.CommonSlideData.InsertBefore(newBackground, _PresentationSlide.CommonSlideData.ShapeTree);
            }
        }
        
        public IEnumerable<Dictionary<string, string>> GetComments()
        {
            List<Dictionary<string, string>> comments = new List<Dictionary<string, string>>();

            if (_CommentPart != null)
            {
                var commentList = _CommentPart.CommentList;
                // Extract comment authors
                foreach (var comment in commentList.Elements<Comment>())
                {
                    Dictionary<string, string> CommentProperties = new Dictionary<string, string>
                    {
                        { "Index", comment.Index},
                        { "Text", comment.InnerText },
                        { "AuthorId", comment.AuthorId },
                        { "DateTime", comment.DateTime },
                        { "XPos", comment.Descendants<P.Position>().FirstOrDefault()?.X ?? 0},
                        { "YPos", comment.Descendants<P.Position>().FirstOrDefault()?.X ?? 0 }
                    };

                    comments.Add(CommentProperties);
                }
            }
            return comments;
        }
        public TextShapeFacade AddTextShape (String text, Int32 fontSize, TextAlignment alignment, Int64 _x, Int64 _y, 
            Int64 width, Int64 height, String fontFamily, String textColor, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithText(text)
                .WithFontSize(fontSize)
                .WithFontFamily(fontFamily)
                .WithTextColor(textColor)
                .WithBackgroundColor(backgroundColor)
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            //_TextShapeFacades.Add(textShapeFacade);
            return textShapeFacade;
        }
        public RectangleShapeFacade DrawRectangle( Int64 _x, Int64 _y,
           Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            RectangleShapeFacade rectangleShapeFacade = new RectangleShapeFacade();

            // Set properties using the provided parameters
            rectangleShapeFacade                
                .WithBackgroundColor(backgroundColor)                
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape rectangleShape = rectangleShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(rectangleShape);
            
            return rectangleShapeFacade;
        }
        public TriangleShapeFacade DrawTriangle(Int64 _x, Int64 _y,
           Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TriangleShapeFacade TriangleShapeFacade = new TriangleShapeFacade();

            // Set properties using the provided parameters
            TriangleShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape TriangleShape = TriangleShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(TriangleShape);

            return TriangleShapeFacade;
        }
        public DiamondShapeFacade DrawDiamond(Int64 _x, Int64 _y,
           Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            DiamondShapeFacade DiamondShapeFacade = new DiamondShapeFacade();

            // Set properties using the provided parameters
            DiamondShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape DiamondShape = DiamondShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(DiamondShape);

            return DiamondShapeFacade;
        }
        public LineFacade DrawLine(Int64 _x, Int64 _y,
           Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            LineFacade _LineFacade = new LineFacade();

            // Set properties using the provided parameters
            _LineFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Line = _LineFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Line);

            return _LineFacade;
        }
        public CurvedLineFacade DrawCurvedLine(Int64 _x, Int64 _y,
           Int64 width, Int64 height)
        {
            // Create an instance of Curved Line
            CurvedLineFacade _CurvedLineFacade = new CurvedLineFacade();

            // Set properties using the provided parameters
            _CurvedLineFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Line = _CurvedLineFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Line);

            return _CurvedLineFacade;
        }
        public ArrowFacade DrawArrow(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            ArrowFacade _ArrowFacade = new ArrowFacade();

            // Set properties using the provided parameters
            _ArrowFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Arrow = _ArrowFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Arrow);

            return _ArrowFacade;
        }
        public DoubleArrowFacade DrawDoubleArrow(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            DoubleArrowFacade _DoubleArrowFacade = new DoubleArrowFacade();

            // Set properties using the provided parameters
            _DoubleArrowFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _DoubleArrow = _DoubleArrowFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_DoubleArrow);

            return _DoubleArrowFacade;
        }
        public DoubleBraceFacade DrawDoubleBrace(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            DoubleBraceFacade _DoubleBraceFacade = new DoubleBraceFacade();

            // Set properties using the provided parameters
            _DoubleBraceFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _DoubleBrace = _DoubleBraceFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_DoubleBrace);

            return _DoubleBraceFacade;
        }
        public DoubleBracketFacade DrawDoubleBracket(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            DoubleBracketFacade _DoubleBracketFacade = new DoubleBracketFacade();

            // Set properties using the provided parameters
            _DoubleBracketFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _DoubleBracket = _DoubleBracketFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_DoubleBracket);

            return _DoubleBracketFacade;
        }
        public PentagonFacade DrawPentagon(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            PentagonFacade _PentagonFacade = new PentagonFacade();

            // Set properties using the provided parameters
            _PentagonFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Pentagon = _PentagonFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Pentagon);

            return _PentagonFacade;
        }

        public PieFacade DrawPie(Int64 _x, Int64 _y,
          Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            PieFacade _PieFacade = new PieFacade();

            // Set properties using the provided parameters
            _PieFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Pie = _PieFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Pie);

            return _PieFacade;
        }
        public HexagonFacade DrawHexagon(Int64 _x, Int64 _y,
                 Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            HexagonFacade _HexagonFacade = new HexagonFacade();

            // Set properties using the provided parameters
            _HexagonFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Hexagon = _HexagonFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Hexagon);

            return _HexagonFacade;
        }
        public TrapezoidFacade DrawTrapezoid(Int64 _x, Int64 _y,
                 Int64 width, Int64 height)
        {
            // Create an instance of TextShapeFacade
            TrapezoidFacade _TrapezoidFacade = new TrapezoidFacade();

            // Set properties using the provided parameters
            _TrapezoidFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Trapezoid = _TrapezoidFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Trapezoid);

            return _TrapezoidFacade;
        }
        public CircleShapeFacade DrawCircle(Int64 _x, Int64 _y,
          Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            CircleShapeFacade circleShapeFacade = new CircleShapeFacade();

            // Set properties using the provided parameters
            circleShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape CircleShape = circleShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(CircleShape);

            return circleShapeFacade;
        }
        public TextShapeFacade AddTextListShape (List<String> textList, ListFacade facade, Int32 fontSize, TextAlignment alignment, Int64 _x, Int64 _y,
            Int64 width, Int64 height, String fontFamily, String textColor, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithFontSize(fontSize)
                .WithFontFamily(fontFamily)
                .WithTextColor(textColor)
                .WithBackgroundColor(backgroundColor)
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateListShape(textList,facade);

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            //_TextShapeFacades.Add(textShapeFacade);
            return textShapeFacade;
        }
        public TextShapeFacade AddTextShape ( List<TextSegmentFacade> textSegmentFacades,TextAlignment alignment, Int64 _x, Int64 _y,
           Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height)
                .WithBackgroundColor(backgroundColor);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateShape(textSegmentFacades);

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            //_TextShapeFacades.Add(textShapeFacade);
            return textShapeFacade;
        }

        public void AddNote(String noteText)
        {
            var relId = _RelationshipId;
           
            NotesSlidePart notesSlidePart1;
            string existingSlideNote = noteText;

            if (_SlidePart.NotesSlidePart != null)
            {
                //Appened new note to existing note.
                existingSlideNote = _SlidePart.NotesSlidePart.NotesSlide.InnerText + "\n" + noteText;
                //var val = (NotesSlidePart)_SlidePart.GetPartById(relId);
                //var val = _SlidePart.NotesSlidePart;
                notesSlidePart1 = _NotesPart;
            }
            else
            {
                //Add a new noteto a slide.                      
                notesSlidePart1 = _SlidePart.AddNewPart<NotesSlidePart>(relId);
            }

            NotesSlide notesSlide = new NotesSlide(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Slide Image Placeholder 1" },
                            new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.SlideImage })),
                        new P.ShapeProperties()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Notes Placeholder 2" },
                            new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(
                                new D.Run(
                                    new D.RunProperties() { Language = "en-US", Dirty = false },
                                    new D.Text() { Text = existingSlideNote }),
                                new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false }))
                            ))),
                new ColorMapOverride(new D.MasterColorMapping()));

            notesSlidePart1.NotesSlide = notesSlide;
        }
        public void RemoveNote()
        {
            if (_SlidePart.NotesSlidePart != null)
            {
                // Clear the existing notes.
                _SlidePart.DeletePart(_SlidePart.NotesSlidePart);
            }
        }
        public void  AddImage (ImageFacade picture )
        {
            _PresentationSlide.CommonSlideData.ShapeTree.Append(picture.Image);
        }
       
        public void Update ()
        {
            this.SetSlideBackground(_BackgroundColor);
        }

    }
}
