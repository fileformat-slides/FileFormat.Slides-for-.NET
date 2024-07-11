using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using FileFormat.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace FileFormat.Slides
{
    /// <summary>
    /// Represents the slide object within a presentatction
    /// </summary>
    public class Slide
    {
        private SlideFacade _SlideFacade;
        private String _RelationshipId;
        private int _SlideIndex;
        private List<TextShape> _TextShapes;
        private List<Rectangle> _Rectangles;
        private List<Circle> _Circles;
        private List<Image> _Images;
        private List<Table> _Tables;
        private static String _BackgroundColor = null;
        private CommentFacade _CommentFacade=null;
        private Presentation _SlidePresentation;
        private int _CommentIndex = 0;

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
        /// Property to set background color of a slide.
        /// </summary>
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        /// <summary>
        /// Property to get or set the list of tables
        /// </summary>
        public List<Table> Tables { get => _Tables; set => _Tables = value; }
        /// <summary>
        /// Property to get or set the relative presentation instance
        /// </summary>
        public Presentation SlidePresentation { get => _SlidePresentation; set => _SlidePresentation = value; }
        /// <summary>
        /// Property to get or set the list of Rectangles.
        /// </summary>
        public List<Rectangle> Rectangles { get => _Rectangles; set => _Rectangles = value; }
        /// <summary>
        /// Property to get or set list of circles.
        /// </summary>
        public List<Circle> Circles { get => _Circles; set => _Circles = value; }


        /// <summary>
        /// Constructor for the Slide class.
        /// </summary>
        /// <remarks>
        ///  it intializes the Slide Facade set the slide index and intializes the lists of text shapes and images.
        /// </remarks>
        public Slide ()
        {
            try
            {
                _SlideIndex = Utility.SlideNextIndex;
                Utility.SlideNextIndex += 1;
                _SlideFacade = new SlideFacade(true);
                _SlideFacade.SlideIndex = _SlideIndex;
                _RelationshipId = _SlideFacade.RelationshipId;
                _TextShapes = new List<TextShape>();
                _Rectangles = new List<Rectangle>();
                _Circles = new List<Circle>();
                _Images = new List<Image>();
                _Tables = new List<Table>();
                _SlideFacade.BackgroundColor = _BackgroundColor;
                _CommentFacade = new CommentFacade();
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Initializing slide");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Contructor which accepts bool value 
        /// </summary>
        /// <param name="isNewSlide"></param>
        public Slide (bool isNewSlide)
        {
            if (isNewSlide)
                _SlideFacade = new SlideFacade(true);
            else
                _SlideFacade = new SlideFacade(false);


            _RelationshipId = _SlideFacade.RelationshipId;
            _TextShapes = new List<TextShape>();
            _Rectangles = new List<Rectangle>();
            _Circles = new List<Circle>();
            _Images = new List<Image>();
            _Tables = new List<Table>();
            _CommentFacade = new CommentFacade();
        }
        /// <summary>
        /// Method to add a text shape in a slide.
        /// </summary>
        /// <param name="textShape">An object of TextShape class.</param>
        public void AddTextShapes (TextShape textShape)
        {
            try
            {
                if (textShape.TextList == null)
                {
                    textShape.Facade = _SlideFacade.AddTextShape(textShape.Text, textShape.FontSize, TextAlignment.Center,
                        Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                        , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.FontFamily,
                        textShape.TextColor, textShape.BackgroundColor);
                }
                else
                {
                    textShape.Facade = _SlideFacade.AddTextListShape(textShape.TextList.ListItems, textShape.TextList.Facade, textShape.FontSize, TextAlignment.Center,
                        Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                        , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.FontFamily,
                        textShape.TextColor, textShape.BackgroundColor);
                }
                textShape.ShapeIndex = _TextShapes.Count + 1;
                _TextShapes.Add(textShape);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding text shape");
                throw new Common.FileFormatException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to draw a rectangular shape
        /// </summary>
        /// <param name="rect"></param>
        /// <exception cref="Common.FileFormatException"></exception>
        public void DrawRectangle(Rectangle rect)
        {          

            try
            {
                rect.Facade = _SlideFacade.DrawRectangle(Utility.PixelsToEmu(rect.X), Utility.PixelsToEmu(rect.Y), 
                    Utility.PixelsToEmu(rect.Width), Utility.PixelsToEmu(rect.Height), rect.BackgroundColor);
                rect.ShapeIndex = _Rectangles.Count + 1;
                _Rectangles.Add(rect);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding text shape");
                throw new Common.FileFormatException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to draw a circle in a slide 
        /// </summary>
        /// <param name="circle"></param>
        /// <exception cref="Common.FileFormatException"></exception>
        public void DrawCircle(Circle circle)
        {
            try
            {
                circle.Facade = _SlideFacade.DrawCircle(Utility.PixelsToEmu(circle.X), Utility.PixelsToEmu(circle.Y),
                    Utility.PixelsToEmu(circle.Width), Utility.PixelsToEmu(circle.Height), circle.BackgroundColor);
                circle.ShapeIndex = _Rectangles.Count + 1;
                _Circles.Add(circle);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding text shape");
                throw new Common.FileFormatException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to add/update note to a slide
        /// </summary>
        /// <param name="noteText">Text you want to add as note</param>
        public void AddNote(String noteText)
        {
            _SlideFacade.AddNote(noteText);
        }
        /// <summary>
        /// Method to remove Notes of a slide
        /// </summary>
        public void RemoveNote()
        {
            _SlideFacade.RemoveNote();
        }
        public void AddTextShapes (TextShape textShape, List<TextSegment> textSegments)
        {
            try
            {
                List<TextSegmentFacade> textSegmentFacades = new List<TextSegmentFacade>();
                foreach (var ts in textSegments)
                {
                    textSegmentFacades.Add(ts.Facade);
                }
                textShape.Facade = _SlideFacade.AddTextShape(textSegmentFacades, TextAlignment.Center,
                    Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                    , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.BackgroundColor);
                textShape.ShapeIndex = _TextShapes.Count + 1;
                _TextShapes.Add(textShape);
            }
            catch (Exception ex)
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
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding image");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Method to add comments to a slide. 
        /// </summary>
        /// <param name="comment">An object of Comment class</param>
        public void AddComment(Comment comment)
        {
            try
            {
                comment.Facade = _CommentFacade;
                if (_CommentIndex == 0)
                {                    
                    comment.Facade.SetAssociatedSlidePart(_SlideFacade.SlidePart, _SlidePresentation.Facade.CommentAuthorPart);
                }
                UInt32Value id = new UInt32Value { Value = (uint)comment.AuthorId };
                comment.Facade.GenerateComment(id,comment.Text, comment.InsertedAt, comment.X, comment.Y);
                _CommentIndex += 1;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding Comment");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Method to get the list of comments.
        /// </summary>
        /// <returns></returns>
        public List<Comment> GetComments()
        {
            var comments= new List<Comment>();
            var facadeComments= _SlideFacade.GetComments();
            foreach (var facadeComment in facadeComments)
            {
                Comment comment = new Comment();
                comment.Text = facadeComment["Text"];
                comment.AuthorId = Convert.ToInt32(facadeComment["AuthorId"]);
                comment.CommentIndex= Convert.ToInt32(facadeComment["Index"]);
                comment.InsertedAt = Convert.ToDateTime(facadeComment["DateTime"]);
                comment.X = Convert.ToInt64(facadeComment["XPos"]);
                comment.Y = Convert.ToInt64(facadeComment["YPos"]);
                comment.Facade = new CommentFacade();
                comment.Facade.CommentPart = _SlideFacade.CommentPart;

                comments.Add(comment);
            }
            return comments;
        }
        /// <summary>
        /// Method to add table to a slide. 
        /// </summary>
        /// <param name="table">An object of Table class</param>
        public void AddTable (Table table)
        {
            try
            {
                table.Facade = new TableFacade();
                table.Facade.X = Utility.PixelsToEmu(table.X);
                table.Facade.Y = Utility.PixelsToEmu(table.Y);
                table.Facade.Width = Utility.PixelsToEmu(table.Width);
                table.Facade.Height = Utility.PixelsToEmu(table.Height);
                if (table.Theme == null)
                {
                    table.Theme = Table.TableStyle.LightStyle1;
                }
                
                table.Facade.GenerateTable(_SlideFacade.SlidePart, table.GetDataTable());
                table.TableIndex = _Tables.Count + 1;
                _Tables.Add(table);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding table");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }        

        /// <summary>
        /// Get text shapes by searching a text term.
        /// </summary>
        /// <param name="text">Search term as string</param>
        /// <returns></returns>
        public List<TextShape> GetTextShapesByText (String text)
        {
            try
            {
                List<TextShape> shapes = TextShapes.Where(shape => shape.Text.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                return shapes;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Getting Shapes");
                throw new Common.FileFormatException(errorMessage, ex);
            }
        }
        
        /// <summary>
        /// Method to update a slide properties e.g. background color.
        /// </summary>
        public void Update ()
        {
            _SlideFacade.BackgroundColor = _BackgroundColor;
            _SlideFacade.Update();
        }

    }
}
