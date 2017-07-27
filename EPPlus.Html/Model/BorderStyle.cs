﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
  
    public class BorderStyle: IBorderStyle
    {

        /// <summary>
        /// Width of the border in pixels
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Color of the border
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Style of the line (e.g. solid, dashed)
        /// </summary>
        public string LineStyle { get; set; }

    }
}
