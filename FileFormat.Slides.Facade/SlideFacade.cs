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

        private List<ImageFacade> _ImagesFacade = null;

        private List<TableFacade> _TableFacades = null;

        private CommentAuthorsPart _CommentAuthorPart;

        private SlideCommentsPart _CommentPart;

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
