# FileFormat.Slides for .NET | Free C# PowerPoint API
[FileFormat.Slides for .NET](https://github.com/fileformat-slides/FileFormat.Slides-for-.NET) - An open-source library that can help beginners create, open, and edit PowerPoint files.

# C# PowerPoint API System Outlines
This Documentation explains the internal structure of our PowerPoint Presentation Management C# API system. Despite its complexity, we've ensured the public APIs are user-friendly, providing an easy experience for manipulating PowerPoint presentations.

For a more detailed understanding of our system architecture, design patterns, and public interfaces, please visit the [Articles Section](https://fileformat-slides.github.io/FileFormat.Slides-for-.NET/docs/index.html).

## FileFormat.Slides Namespace

### Presentation Class
- The primary class responsible for creating, loading, and modifying presentations.

### Slide Class
- This class represents the slides of a presentation. It deals with elements creation, updation, retrieval and deletion operations within a slide.

### TextShape Class 
- This class is responsible to manage the text shapes within a slide.
- It allows add, update, retrieve and removing of a text shape.
- It allows to set text, x and y coordinates, width, height, font size, font color, font family, text alignment of a text shape.

### Image Class
- This class is providing the functions to deal with Image within a slide.
- It allows add, update, retrieve and removing of an image.

### StyledList Class
- This class facilitates the addition of numbered or bulleted lists.
- It enables easy updates and removals of list items.
- Users can change the styling of entire lists or individual items.
- Complete lists can be updated or removed.
- It allows conversion between numbered and bulleted lists.

### Table Class
- Enables users to add tables to PPT/PPTX slides.
- Supports styling at the table, row, and cell levels.

## FileFormat.Slides.Common Namespace
- This namespace contains all classes, enums or methods for common use.

### Utility Class
- This class provides essential static methods for generating unique relationship IDs, obtaining random slide IDs, and converting measurements.

### Colors Class
- This static class provides static properties with color codes, simplifying consistent color selection in C# applications.

## FileFormat.Slides.Facade
- Contains facade classes


## API Reference
- [API Reference](https://fileformat-slides.github.io/FileFormat.Slides-for-.NET/api/FileFormat.Slides.html) - In-depth information about public interfaces and usage.

## Technical Docs
- [Articles](https://fileformat-slides.github.io/FileFormat.Slides-for-.NET/docs/index.html) - Comprehensive insights into the system architecture, design patterns, and API usage in different scenarios.

# Installation
- Install-Package FileFormat.Slides

# System Requirements
- .NET Core 3.1 and above
