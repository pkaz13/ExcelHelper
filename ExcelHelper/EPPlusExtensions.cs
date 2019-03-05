using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    static class EPPlusExtensions
    {
        /// <summary>
        /// Set default table style in LoadFromCollection method, third parameter.
        /// </summary>
        public static ExcelRangeBase LoadFromCollectionFiltered<T>(this ExcelRangeBase @this, IEnumerable<T> collection)
        {
            MemberInfo[] membersToInclude = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(x => !Attribute.IsDefined(x, typeof(EPPlusIgnore))).ToArray();
            return @this.LoadFromCollection(collection, true, OfficeOpenXml.Table.TableStyles.Dark2, BindingFlags.Instance | BindingFlags.Public, membersToInclude);
        }

        /// <summary>
        /// Converts columns from workseet to specific data types.
        /// </summary>
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {
            Func<CustomAttributeData, bool> columnOnly = x => x.AttributeType == typeof(ExcelColumn);

            var columns = typeof(T).GetProperties().Where(x => x.CustomAttributes.Any(columnOnly)).Select(x => new
            {
                Property = x,
                Column = x.GetCustomAttributes<ExcelColumn>().First().ColumnIndex
            }).ToList();

            var rows = worksheet.Cells.Select(x => x.Start.Row).Distinct().OrderBy(x => x);

            var collection = rows.Skip(1).Select(row =>
            {
                var tnew = new T();
                columns.ForEach(col =>
                {
                    var val = worksheet.Cells[row, col.Column];
                    if (val.Value == null)
                    {
                        col.Property.SetValue(tnew, null);
                        return;
                    }
                    if (col.Property.PropertyType == typeof(Int32))
                    {
                        col.Property.SetValue(tnew, val.GetValue<int>());
                        return;
                    }
                    if (col.Property.PropertyType == typeof(double))
                    {
                        col.Property.SetValue(tnew, val.GetValue<double>());
                        return;
                    }
                    if (col.Property.PropertyType == typeof(DateTime))
                    {
                        col.Property.SetValue(tnew, val.GetValue<DateTime>());
                        return;
                    }
                    if (col.Property.PropertyType == typeof(decimal))
                    {
                        col.Property.SetValue(tnew, val.GetValue<decimal>());
                        return;
                    }
                    col.Property.SetValue(tnew, val.GetValue<string>());
                });
                return tnew;
            });
            return collection;
        }
    }
}
