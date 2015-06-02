using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI
{
    public class GridViewSort<T>
    {
        private bool _sortedAscending;
        private string _columnName;

        public GridViewSort(string ColumnName, bool SortedAscending)
        {
            _columnName = ColumnName;
            _sortedAscending = SortedAscending;
        }

        public IEnumerable<T> Sort(IEnumerable<T> Items, string ColumnName)
        {
            if (ColumnName == _columnName && _sortedAscending)
            {
                _sortedAscending = false;
                var test = Items.Select(x => x.GetType().GetProperty(ColumnName).GetValue(x));
                return Items.OrderByDescending(x => x.GetType().GetProperty(ColumnName).GetValue(x));
            }
            else
            {
                _columnName = ColumnName;
                _sortedAscending = true;
                return Items.OrderBy(x => x.GetType().GetProperty(ColumnName).GetValue(x));
            }
        }
    }
}
