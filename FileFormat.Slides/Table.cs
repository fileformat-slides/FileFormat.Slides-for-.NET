using DocumentFormat.OpenXml.Spreadsheet;
using FileFormat.Slides.Common;
using FileFormat.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Data;

namespace FileFormat.Slides
{
    /// <summary>
    /// This class is responsible to create table in a PPT/PPTX presentataion.
    /// </summary>
    public class Table
    {

        private TableFacade _Facade;
        private int _TableIndex = 0;
        private String _Name;
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private List<TableRow> _Rows;
        private List<TableColumn> _Columns;
        private Stylings _TableStylings;



        /// <summary>
        /// Property to get or set the TableFacade instance.
        /// </summary>
        public TableFacade Facade { get => _Facade; set => _Facade = value; }
        /// <summary>
        /// Property to get or set the index of a table within slide.
        /// </summary>
        public int TableIndex { get => _TableIndex; set => _Facade.TableIndex = _TableIndex = value; }
        /// <summary>
        /// Property to get or set the table name within the slide.
        /// </summary>
        public string Name { get => _Name; set => _Name = value; }
        /// <summary>
        /// Property to get or set the X coordinate of a table.
        /// </summary>
        public double X { get => _x; set => _x = value; }
        /// <summary>
        /// Property to get or set the Y coordinate of a table.
        /// </summary>
        public double Y { get => _y; set => _y = value; }
        /// <summary>
        /// Property to get or set the width of a table.
        /// </summary>
        public double Width { get => _Width; set => _Width = value; }
        /// <summary>
        /// Property to get or set the height of a table.
        /// </summary>
        public double Height { get => _Height; set => _Height = value; }

        /// <summary>
        /// Property to get or set the list of rows in the table.
        /// </summary>
        public List<TableRow> Rows { get => _Rows; set => _Rows = value; }

        /// <summary>
        /// Property to get or set the list of columns in the table.
        /// </summary>
        public List<TableColumn> Columns { get => _Columns; set => _Columns = value; }

        /// <summary>
        /// Property to get or set the stylings for the table.
        /// </summary>
        public Stylings TableStylings { get => _TableStylings; set => _TableStylings = value; }

        /// <summary>
        /// Constructor for the Table class. Initializes a new instance of the Table class with empty lists for rows and columns.
        /// </summary>
        public Table()
        {
            _Rows = new List<TableRow>();
            _Columns = new List<TableColumn>();
        }

        /// <summary>
        /// Adds a row to the table.
        /// </summary>
        /// <param name="row">The TableRow object to be added to the table.</param>
        public void AddRow(TableRow row)
        {
            // If table stylings are defined, apply them to the row.
            if (_TableStylings.FontSize > 0)
            {
                row.RowStylings = _TableStylings;
            }

            // Add the row to the table.
            _Rows.Add(row);
        }



    }
}
