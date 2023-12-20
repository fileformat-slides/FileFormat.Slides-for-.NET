using System;
using System.Collections.Generic;
using System.Text;
using FileFormat.Slides.Common.Enumerations;
using FileFormat.Slides.Facade;
using FileFormat.Slides.Common;

namespace FileFormat.Slides
{
    /// <summary>
    ///  This class represents the text list with bullet style.
    /// </summary>
    public class StyledList
    {
        private ListFacade _Facade;
        private List<String> _ListItems;
        public List<string> ListItems { get => _ListItems; set => _ListItems = value; }
        /// <summary>
        /// Property to get the facade of a styled list
        /// </summary>
        public ListFacade Facade { get => _Facade; }
        /// <summary>
        /// Constructor of StyledList class.
        /// </summary>
        public StyledList ()
        {
            _Facade = new ListFacade();
            _ListItems = new List<String>();
        }
        /// <summary>
        /// Method to add list items in styled list.
        /// </summary>
        /// <param name="text">It accepts text as list item</param>
        public void AddListItem (String text)
        {
            _ListItems.Add(text);
        }

       

       
    }
}
