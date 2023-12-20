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
        /*
        Presentation presentation = Presentation.Create("D:\\AsposeSampleResults\\test2.pptx");
        TextShape shape = new TextShape();
        shape.Text = "Title: Here is my first title From FF";
        shape.BackgroundColor = "5f7200";
        shape.FontSize = 80;
        shape.TextColor = "980078";
        shape.FontFamily = "Baguet Script";
        TextShape shape2 = new TextShape();

        shape2.BackgroundColor = "ff7f90";
        List<TextSegment> TextSegments = new List<TextSegment>();
        TextSegments.Add(new TextSegment{Color= "980078", FontSize = 70, FontFamily = "Calibri", Text = "Body:" }.create());
        TextSegments.Add(new TextSegment{ Color = "000000", FontSize = 32, FontFamily = "Baguet Script", Text = " Here is my text Segment" }.create());
        
        shape2.Y = Utility.EmuToPixels(3499619);
        // First slide
        Slide slide = new Slide();
        slide.AddTextShapes(shape);
        slide.AddTextShapes(shape2, TextSegments);
        
        // Adding slides
        presentation.AppendSlide(slide);
        
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
         // Add slide to an existing presentation
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


        /*
       //Update an image in a slide
       Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
       var slides = presentation.GetSlides();
       var slide = slides[0];
       var shape = slide.TextShapes[0];
        shape.BackgroundColor = "7f7f88";
        shape.Text = "Updated 2nd Text";
        shape.Update();
        Slide slide1 = new Slide();
        slide1.BackgroundColor = Colors.Green;
        Image image1 = new Image("D:\\AsposeSampleData\\target.png");
        image1.X = Utility.EmuToPixels(1838700);
        image1.Y = Utility.EmuToPixels(1285962);
        image1.Width = Utility.EmuToPixels(2514600);
        image1.Height = Utility.EmuToPixels(2886075);
        slide1.AddImage(image1);
        
        presentation.AppendSlide(slide1);
        presentation.Save();
        */
        /*
         * Add bulleted list.
        Presentation presentation = Presentation.Create("D:\\AsposeSampleResults\\test2.pptx");
        Slide slide1 = new Slide();
        slide1.BackgroundColor = Colors.Purple;
        TextShape shape1 = new TextShape();
        shape1.FontFamily = "Baguet Script";
        shape1.FontSize = 60;
        shape1.Y = 200.0;
        shape1.TextColor = Colors.Yellow;
        shape1.BackgroundColor = Colors.LimeGreen;

        StyledList list = new StyledList();
        list.AddListItem("Pakistan");
        list.AddListItem("India");
        list.AddListItem("Australia");
        list.AddListItem("England");

        shape1.TextList = list;
        slide1.AddTextShapes(shape1);

        presentation.AppendSlide(slide1);
        presentation.Save();
        */


    }

}
