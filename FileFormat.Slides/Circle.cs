﻿using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using FileFormat.Slides.Facade;
using System;
using System.Collections.Generic;

namespace FileFormat.Slides
{
    /// <summary>
    /// This class represents the circle shape within a slide.
    /// </summary>
    public class Circle
    {
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private CircleShapeFacade _Facade;
        private int _shapeIndex;
        private string _BackgroundColor = null;

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
        /// Property to get or set the CircleShapeFacade.
        /// </summary>
        public CircleShapeFacade Facade { get => _Facade; set => _Facade = value; }

        /// <summary>
        /// Property to get or set the shape index within a slide.
        /// </summary>
        public int ShapeIndex { get => _shapeIndex; set => _shapeIndex = value; }

        /// <summary>
        /// Property to set or get background color of a circle shape.
        /// </summary>
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }

        /// <summary>
        /// Constructor of the Circle class initializes the object of CircleShapeFacade and populates its fields.
        /// </summary>
        public Circle()
        {
            _Facade = new CircleShapeFacade();
            _Facade.ShapeIndex = _shapeIndex;

            _BackgroundColor = "Transparent";
            _x = Utility.EmuToPixels(1349828);
            _y = Utility.EmuToPixels(1999619);
            _Width = Utility.EmuToPixels(6000000);
            _Height = Utility.EmuToPixels(2000000);

            Populate_Facade();
        }

        /// <summary>
        /// Method to update circle shape.
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
        /// Method for getting the list of circle shapes.
        /// </summary>
        /// <param name="CircleFacades">A list of CircleShapeFacade objects.</param>
        /// <returns>A list of Circle objects.</returns>
        public static List<Circle> GetCircles(List<CircleShapeFacade> CircleFacades)
        {
            List<Circle> Circles = new List<Circle>();
            try
            {
                foreach (var facade in CircleFacades)
                {
                    Circle Circle = new Circle
                    {
                        BackgroundColor = facade.BackgroundColor,
                        X = Utility.EmuToPixels(facade.X),
                        Y = Utility.EmuToPixels(facade.Y),
                        Width = Utility.EmuToPixels(facade.Width),
                        Height = Utility.EmuToPixels(facade.Height),
                        Facade = facade,
                        ShapeIndex = facade.ShapeIndex
                    };

                    Circles.Add(Circle);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = Common.FileFormatException.ConstructMessage(ex, "Getting Circle Shapes");
                throw new Common.FileFormatException(errorMessage, ex);
            }

            return Circles;
        }

        /// <summary>
        /// Method to remove the circle shape from a slide.
        /// </summary>
        public void Remove()
        {
            _Facade.RemoveShape(this.Facade.CircleShape);
        }
    }
}
