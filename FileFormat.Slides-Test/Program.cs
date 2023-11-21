using DocumentFormat.OpenXml;
using PKG = DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using GeneratedCode;
using FileFormat.Slides;
using System.Collections.Generic;
using FileFormat.Slides.Common;

class Program
{
    static void Main ()
    { 
        /* Create new Presentation
         Presentation presentation = Presentation.Create("D:\\AsposeSampleResults\\test2.pptx");
         TextShape shape = new TextShape();
         shape.Text = "Title: Here is my first title From FF";
         shape.TextColor = "980078";
         shape.FontFamily = "Baguet Script";
         TextShape shape2 = new TextShape();
         shape2.Text = "Body : Here is my first title From FF";
         shape2.FontFamily = "BIZ UDGothic";
         shape2.FontSize = 3000;
         shape2.Y = Utility.EmuToPixels(2499619);
         // First slide
         Slide slide = new Slide();
         slide.AddTextShapes(shape);
         slide.AddTextShapes(shape2);
         // 2nd slide
         Slide slide1 = new Slide();
         slide1.AddTextShapes(shape);
         slide1.AddTextShapes(shape2);
         // Adding slides
         presentation.AppendSlide(slide);
         presentation.AppendSlide(slide1);
         presentation.Save();*/
        
        /* Open and update a PPTX file
        Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        var slides = presentation.GetSlides();
        var slide = slides[3];
        List<TextShape> shapes = slide.GetTextShapesByText("PRESENTATION");
        var shape = slide.TextShapes[1];
        //shape.X = 100000;
        //shape.Y = 100000;
        shape.Alignment= FileFormat.Slides.Common.Enumerations.TextAlignment.Left;
        shape.Text = "Muhammad Umar";
        shape.Update();
        presentation.Save();
        */

        /* 
         * Remove a slide from presentation.
        Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        var confirmation = presentation.RemoveSlide(0);
        Console.WriteLine(confirmation);
        presentation.Save();
        */
        
        /*
         * Remove text shape from a slide
        Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        var slides = presentation.GetSlides();
        var slide = slides[0];
        var shape = slide.TextShapes[0];
        shape.Remove();
        presentation.Save();
        */


        /*
         * Add slide to an existing presentation
         Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
         TextShape shape1 = new TextShape();
         shape1.Text = "Body : Here is my first title From FF";
         shape1.FontFamily = "Baguet Script";
         shape1.TextColor = Colors.Olive;
         shape1.FontSize = 45;
         shape1.Y = 10.0;
         // First slide
         Slide slide = new Slide();
         Image image1 = new Image("D:\\AsposeSampleData\\target.png");
         image1.X = Utility.EmuToPixels(1838700);
         image1.Y = Utility.EmuToPixels(1285962);
         image1.Width = Utility.EmuToPixels(2514600);
         image1.Height = Utility.EmuToPixels(2886075);
         slide.AddImage(image1);
         slide.AddTextShapes(shape1);        
         presentation.AppendSlide(slide);
         presentation.Save();
        */

        /*
         * Update an image in a slide
        Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        var slides = presentation.GetSlides();
        var slide = slides[0];
        List<Image> images = slide.Images;
        var image = slide.Images[0];
        image.Width = 300.0;
        image.Height = 300.0;
        image.Update();
        presentation.Save();*/

        /*
         * Extract and save images of an existing PPTX
        Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        presentation.ExtractAndSaveImages("testing images"); */




    }

}
