using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    /// <summary>
    /// Add this attribute to model property, which you want exclude from Excel file.
    /// </summary>
    public class EPPlusIgnore : Attribute
    {
    }
}
