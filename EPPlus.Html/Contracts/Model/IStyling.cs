using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public enum HAlign { Left, Center, Right }
    public enum VAlign { Top, Middle, Bottom }

    public interface IStyling
    {

        /// <summary>
        /// Background color
        /// </summary>
        string Background { get; set; }

        /// <summary>
        /// True if the text is bold
        /// </summary>
        bool IsBold { get; set; }

        /// <summary>
        /// Number of pixels for the top border
        /// </summary>
        IBorderStyle BorderTop { get; set; }

        /// <summary>
        /// Number of pixels for the left border
        /// </summary>
        IBorderStyle BorderLeft { get; set; }

        /// <summary>
        /// Number of pixels for the right border
        /// </summary>
        IBorderStyle BorderRight { get; set; }

        /// <summary>
        /// Number of pixels for the bottom border
        /// </summary>
        IBorderStyle BorderBottom { get; set; }

        /// <summary>
        /// Height in pixels
        /// </summary>
        double Height { get; set; }

        /// <summary>
        /// Width in pixels
        /// </summary>
        double Width { get; set; }

        
        /// <summary>
        /// Text color
        /// </summary>
        string Color { get; set; }

        /// <summary>
        /// Font family
        /// </summary>
        string FontFamily { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        double FontSize { get; set; }


        /// <summary>
        /// Number of columns to take up
        /// </summary>
        int Colspan { get; set; }

        /// <summary>
        /// Number of rows to take up
        /// </summary>
        int Rowspan { get; set; }

        /// <summary>
        /// Horizontal alignment
        /// </summary>
        HAlign HAlign { get; set; }

        /// <summary>
        /// Vertical alignment
        /// </summary>
        VAlign VAlign { get; set; }
    }
}
