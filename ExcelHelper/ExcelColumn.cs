using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class ExcelColumn : Attribute
    {
        public int ColumnIndex { get; set; }

        public ExcelColumn(int column)
        {
            ColumnIndex = column;
        }
    }
}
