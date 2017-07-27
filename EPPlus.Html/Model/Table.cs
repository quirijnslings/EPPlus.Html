using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Model
{
    public class Table : ITable
    {
        public Table()
        {
            Styling = new Styling();
            Rows = new List<IRow>();
        }
        public IStyling Styling { get; set; }
        public IList<IRow> Rows { get; set; }
    }
}
