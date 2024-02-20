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
        private List<Image> _Images;
        private List<Table> _Tables;
        private static String _BackgroundColor = null;

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
        public List<Table> Tables { get => _Tables; set => _Tables = value; }


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
                _Images = new List<Image>();
                _Tables = new List<Table>();
                _SlideFacade.BackgroundColor = _BackgroundColor;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Initialing slide");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        public Slide (bool isNewSlide)
        {
            if (isNewSlide)
                _SlideFacade = new SlideFacade(true);
            else
                _SlideFacade = new SlideFacade(false);


            _RelationshipId = _SlideFacade.RelationshipId;
            _TextShapes = new List<TextShape>();
            _Images = new List<Image>();
            _Tables = new List<Table>();
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
        /// Method to add images to a slide. 
        /// </summary>
        /// <param name="image">An object of Image class</param>
        public void AddTable (Table table)
        {
            try
            {
                table.Facade = new TableFacade();
                table.Facade.X = Utility.PixelsToEmu(table.X);
                table.Facade.Y = Utility.PixelsToEmu(table.Y);
                table.Facade.Width = Utility.PixelsToEmu(table.Width);
                table.Facade.Height = Utility.PixelsToEmu(table.Height);
                
                table.Facade.GenerateTable(_SlideFacade.SlidePart, GetDataTable(table));
                table.TableIndex = _Tables.Count + 1;
                _Tables.Add(table);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Adding table");
                throw new Common.FileFormatException(errorMessage, ex);
            }

        }
        private DataTable GetDataTable (Table table)
        {
            DataTable dtable = new DataTable();

            // Adding columns based on TableColumn information
            foreach (TableColumn column in table.Columns)
            {
                dtable.Columns.Add(column.Name, typeof(string));
            }

            // Adding rows based on TableRow and TableCell information
            foreach (TableRow row in table.Rows)
            {
                DataRow dataRow = dtable.NewRow();

                // Assuming each TableCell in the TableRow corresponds to a column in the DataTable
                foreach (TableCell cell in row.Cells)
                {
                    // Find the corresponding column by matching the cell's position
                    // Assuming the order of columns in _Columns corresponds to the order of cells in _Cells
                    int columnIndex = table.Columns.FindIndex(col => col.Name == cell.ID);

                    if (columnIndex >= 0)
                    {
                        // Add cell value to the corresponding column in the DataRow
                        dataRow[columnIndex] = cell.Text;
                        string stylingInfo = Utility.SerializeStyling(cell.CellStylings);
                        dataRow[columnIndex] += ";" + stylingInfo;
                    }
                    else
                    {
                        // Handle the case where the column for the cell is not found
                        // You may want to log a warning or handle it based on your requirements
                        Console.WriteLine($"Column for cell FontFamily {cell.FontFamily} not found in the table.");
                    }
                }

                // Add the populated DataRow to the DataTable
                dtable.Rows.Add(dataRow);
            }

            return dtable;
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
