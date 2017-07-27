using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public enum RowType {  Header, Normal, Total }
    public interface IRow
    {
        RowType RowType { get; set; }
        IList<ICell> Cells { get; set; }
        IStyling Styling { get; set; }
    }
}
