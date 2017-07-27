﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{

    public class Cell : ICell
    {
        /// <summary>
        /// True if cell is merged with cell to the top or to the left
        /// </summary>
        public bool MergedWithTopOrLeft { get; set; }

        /// <summary>
        /// Styling for this cell
        /// </summary>
        public IStyling Styling { get; set; }

        /// <summary>
        /// Text in the cell (or another type of value converted to text)
        /// </summary>
        public string Text { get; set; }
    }
}