using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public class Row : IRow
    {
        public Row()
        {
            Styling = new Styling();
            RowType = RowType.Normal;
            Cells = new List<ICell>();
        }
        public RowType RowType { get; set; }
        public IList<ICell> Cells { get; set; }
        public IStyling Styling { get; set; }
    }
}
