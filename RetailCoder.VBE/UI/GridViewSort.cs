using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI
{
    public class GridViewSort<T>
    {
        public bool SortedAscending { get; private set; }
        public string ColumnName { get; private set; }

        public GridViewSort(string columnName, bool sortedAscending)
        {
            ColumnName = columnName;
            SortedAscending = sortedAscending;
        }

        public IEnumerable<T> Sort(IEnumerable<T> items, string columnName)
        {
            if (columnName == ColumnName && SortedAscending)
            {
                SortedAscending = false;
                return items.OrderByDescending(x => x.GetType().GetProperty(columnName).GetValue(x));
            }
            else
            {
                ColumnName = columnName;
                SortedAscending = true;
                return items.OrderBy(x => x.GetType().GetProperty(columnName).GetValue(x));
            }
        }

        public IEnumerable<T> Sort(IEnumerable<T> items, string columnName, bool sortAscending)
        {
            SortedAscending = sortAscending;
            ColumnName = columnName;

            return sortAscending
                ? items.OrderBy(x => x.GetType().GetProperty(columnName).GetValue(x))
                : items.OrderByDescending(x => x.GetType().GetProperty(columnName).GetValue(x));
        }
    }
}
