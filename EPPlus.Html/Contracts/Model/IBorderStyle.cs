using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
  
    public interface IBorderStyle
    {

        /// <summary>
        /// Width of the border in pixels
        /// </summary>
        int Width { get; set; }

        /// <summary>
        /// Color of the border
        /// </summary>
        string Color { get; set; }

        /// <summary>
        /// Style of the line (e.g. solid, dashed)
        /// </summary>
        string LineStyle { get; set; }

    }
}
