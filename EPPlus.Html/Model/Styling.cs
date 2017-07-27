using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public class Styling : IStyling
    {

        /// <summary>
        /// Background color
        /// </summary>
        public string Background { get; set; }

        /// <summary>
        /// True if the text is bold
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// Number of pixels for the top border
        /// </summary>
        public IBorderStyle BorderTop { get; set; }

        /// <summary>
        /// Number of pixels for the left border
        /// </summary>
        public IBorderStyle BorderLeft { get; set; }

        /// <summary>
        /// Number of pixels for the right border
        /// </summary>
        public IBorderStyle BorderRight { get; set; }

        /// <summary>
        /// Number of pixels for the bottom border
        /// </summary>
        public IBorderStyle BorderBottom { get; set; }

        /// <summary>
        /// Height in pixels
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Width in pixels
        /// </summary>
        public double Width { get; set; }

       
        /// <summary>
        /// Text color
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Font family
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        public double FontSize { get; set; }
        
        /// <summary>
        /// Number of columns to take up
        /// </summary>
        public int Colspan { get; set; }

        /// <summary>
        /// Number of rows to take up
        /// </summary>
        public int Rowspan { get; set; }

        /// <summary>
        /// Horizontal alignment
        /// </summary>
        public HAlign HAlign { get; set; }

        /// <summary>
        /// Vertical alignment
        /// </summary>
        public VAlign VAlign { get; set; }

    }
}
