using System;
using System.Collections.Generic;
using System.Text;
using FileFormat.Slides.Facade;
using FileFormat.Slides.Common.Enumerations;
using System.Linq;
using FileFormat.Slides.Common;

namespace FileFormat.Slides
{
    /// <summary>
    /// Represents the slide object within a presentatction
    /// </summary>
    public class Slide
    {
        private  SlideFacade _SlideFacade;
        private String _RelationshipId;
        private int _SlideIndex;
        private List<TextShape> _TextShapes;       
        private List<Image> _Images;

        /// <summary>
        /// Property for respective Slide Facade.
        /// </summary>
        public SlideFacade SlideFacade { get => _SlideFacade; set => _SlideFacade = value; }
        /// <summary>
        /// Property contains the list of all text shapes.
        /// </summary>
        public List<TextShape> TextShapes { get => _TextShapes; set => _TextShapes = value; }
        /// <summary>
        /// Property for the relationship Id.
        /// </summary>
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }

        /// <summary>
        /// Property to hold the index of the slide.
        /// </summary>
        public int SlideIndex { get => _SlideIndex; set => _SlideIndex = value; }
        /// <summary>
        /// Property contains the list of all images within a slide.
        /// </summary>
        public List<Image> Images { get => _Images; set => _Images = value; }


        /// <summary>
        /// Constructor for the Slide class.
        /// </summary>
        /// <remarks>
        ///  it intializes the Slide Facade set the slide index and intializes the lists of text shapes and images.
        /// </remarks>
        public Slide () {
            try
            {
                _SlideIndex = Utility.SlideNextIndex;
                Utility.SlideNextIndex += 1;
                _SlideFacade = new SlideFacade(true);
                _SlideFacade.SlideIndex = _SlideIndex;
                _RelationshipId = _SlideFacade.RelationshipId;
                _TextShapes = new List<TextShape>();
                _Images = new List<Image>();
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Initialing slide");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        public Slide (bool isNewSlide)
        {
            if(isNewSlide)
                _SlideFacade = new SlideFacade(true);
            else
                _SlideFacade = new SlideFacade(false);

            _RelationshipId = _SlideFacade.RelationshipId;
            _TextShapes = new List<TextShape>();
            _Images = new List<Image>();
        }
        /// <summary>
        /// Method to add a text shape in a slide.
        /// </summary>
        /// <param name="textShape">An object of TextShape class.</param>
        public void AddTextShapes(TextShape textShape)
        {
            try
            {
                textShape.Facade = _SlideFacade.AddTextShape(textShape.Text, textShape.FontSize, TextAlignment.Center,
                    Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                    , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.FontFamily, textShape.TextColor);
                textShape.ShapeIndex = _TextShapes.Count + 1;
                _TextShapes.Add(textShape);
            }catch(Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding text shape");
                throw new Common.FileFormatException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to add images to a slide. 
        /// </summary>
        /// <param name="image">An object of Image class</param>
        public void AddImage (Image image)
        {
            try
            {
                image.Facade = new ImageFacade();
                image.Facade.X = Utility.PixelsToEmu(image.X);
                image.Facade.Y = Utility.PixelsToEmu(image.Y);
                image.Facade.Width = Utility.PixelsToEmu(image.Width);
                image.Facade.Height = Utility.PixelsToEmu(image.Height);
                image.Facade.createImage(image.ImagePath, _SlideFacade.SlidePart);
                _SlideFacade.AddImage(image.Facade);
                image.ImageIndex = _Images.Count + 1;
                _Images.Add(image);
            }catch(Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding image");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Get text shapes by searching a text term.
        /// </summary>
        /// <param name="text">Search term as string</param>
        /// <returns></returns>
        public List<TextShape> GetTextShapesByText(String text) 
        {
            try
            {
                List<TextShape> shapes = TextShapes.Where(shape => shape.Text.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0).ToList();         
                return shapes;  
            }catch(Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Getting Shapes");
                throw new Common.FileFormatException(errorMessage, ex);               
            }
        }
      


    }
}
