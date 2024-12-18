﻿using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using FileFormat.Slides.Facade;
using System;
using System.Collections.Generic;

namespace FileFormat.Slides
{
    /// <summary>
    /// This class represents the rectangle shape within a slide.
    /// </summary>
    public class DoubleBrace
    {
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private DoubleBraceFacade _Facade;
        private int _shapeIndex;
        private string _BackgroundColor = null;
        private AnimationType _Animation = AnimationType.None;
        /// <summary>
        /// Property to get or set X coordinate of the shape.
        /// </summary>
        public double X { get => _x; set => _x = value; }

        /// <summary>
        /// Property to get or set Y coordinate of the shape.
        /// </summary>
        public double Y { get => _y; set => _y = value; }

        /// <summary>
        /// Property to get or set width of the shape.
        /// </summary>
        public double Width { get => _Width; set => _Width = value; }

        /// <summary>
        /// Property to get or set height of the shape.
        /// </summary>
        public double Height { get => _Height; set => _Height = value; }

        /// <summary>
        /// Property to get or set the DoubleBraceFacade.
        /// </summary>
        public DoubleBraceFacade Facade { get => _Facade; set => _Facade = value; }

        /// <summary>
        /// Property to get or set the shape index within a slide.
        /// </summary>
        public int ShapeIndex { get => _shapeIndex; set => _shapeIndex = value; }

        /// <summary>
        /// Property to set or get background color of a rectangle shape.
        /// </summary>
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        /// <summary>
        /// Property to set animation
        /// </summary>
        public AnimationType Animation { get => _Animation; set => _Animation = value; }
        /// <summary>
        /// Constructor of the DoubleBrace class initializes the object of DoubleBraceFacade and populates its fields.
        /// </summary>
        public DoubleBrace()
        {
            _Facade = new DoubleBraceFacade();
            _Facade.ShapeIndex = _shapeIndex;

            _BackgroundColor = "Black";
            _x = Utility.EmuToPixels(1349828);
            _y = Utility.EmuToPixels(1999619);
            _Width = Utility.EmuToPixels(6000000);
            _Height = Utility.EmuToPixels(2000000);

            Populate_Facade();
        }

        /// <summary>
        /// Method to update rectangle shape.
        /// </summary>
        public void Update()
        {
            Populate_Facade();
            _Facade.UpdateShape();
        }

        /// <summary>
        /// Method to populate the fields of the respective facade.
        /// </summary>
        private void Populate_Facade()
        {
            _Facade.BackgroundColor = _BackgroundColor;
            _Facade.X = Utility.PixelsToEmu(_x);
            _Facade.Y = Utility.PixelsToEmu(_y);
            _Facade.Width = Utility.PixelsToEmu(_Width);
            _Facade.Height = Utility.PixelsToEmu(_Height);
        }

        /// <summary>
        /// Method for getting the list of rectangle shapes.
        /// </summary>
        /// <param name="DoubleBraceFacades">A list of DoubleBraceFacade objects.</param>
        /// <returns>A list of DoubleBrace objects.</returns>
        public static List<DoubleBrace> GetDoubleBraces(List<DoubleBraceFacade> DoubleBraceFacades)
        {
            List<DoubleBrace> DoubleBraces = new List<DoubleBrace>();
            try
            {
                foreach (var facade in DoubleBraceFacades)
                {
                    DoubleBrace DoubleBrace = new DoubleBrace
                    {
                        BackgroundColor = facade.BackgroundColor,
                        X = Utility.EmuToPixels(facade.X),
                        Y = Utility.EmuToPixels(facade.Y),
                        Width = Utility.EmuToPixels(facade.Width),
                        Height = Utility.EmuToPixels(facade.Height),
                        Facade = facade,
                        ShapeIndex = facade.ShapeIndex
                    };

                    DoubleBraces.Add(DoubleBrace);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Getting DoubleBrace Shapes");
                throw new Common.FileFormatException(errorMessage, ex);
            }

            return DoubleBraces;
        }

        /// <summary>
        /// Method to remove the rectangle shape from a slide.
        /// </summary>
        public void Remove()
        {
            _Facade.RemoveShape(this.Facade.DoubleBraceShape);
        }
    }
}
