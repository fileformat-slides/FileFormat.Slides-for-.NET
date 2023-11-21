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

        private List<TextShapeFacade> _TextShapeFacades = null;

        private List<ImageFacade> _ImagesFacade = null;
        public Slide PresentationSlide { get => _PresentationSlide; set => _PresentationSlide = value; }
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }
        public SlidePart SlidePart { get => _SlidePart; set => _SlidePart = value; }
        public List<TextShapeFacade> TextShapeFacades { get => _TextShapeFacades; set => _TextShapeFacades = value; }
        public int SlideIndex { get => _SlideIndex; set => _SlideIndex = value; }
        public List<ImageFacade> ImagesFacade { get => _ImagesFacade; set => _ImagesFacade = value; }

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
        public TextShapeFacade AddTextShape (String text, Int32 fontSize, TextAlignment alignment, Int64 _x, Int64 _y, 
            Int64 width, Int64 height, String fontFamily, String textColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithText(text)
                .WithFontSize(fontSize)
                .WithFontFamily(fontFamily)
                .WithTextColor(textColor)
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

        public void  AddImage (ImageFacade picture )
        {
            _PresentationSlide.CommonSlideData.ShapeTree.Append(picture.Image);
        }

    }
}
