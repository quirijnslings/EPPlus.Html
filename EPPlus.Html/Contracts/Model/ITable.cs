using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public interface ITable
    {
        IStyling Styling { get; set; }
        IList<IRow> Rows { get; set; }
    }
}
