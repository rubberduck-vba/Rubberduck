using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI
{
    public class GridViewSort<T>
    {
        private bool _sortedAscending;
        private string _columnName;

        public GridViewSort(string columnName, bool sortedAscending)
        {
            _columnName = columnName;
            _sortedAscending = sortedAscending;
        }

        public IEnumerable<T> Sort(IEnumerable<T> items, string columnName)
        {
            if (columnName == _columnName && _sortedAscending)
            {
                _sortedAscending = false;
                return items.OrderByDescending(x => x.GetType().GetProperty(columnName).GetValue(x));
            }
            else
            {
                _columnName = columnName;
                _sortedAscending = true;
                return items.OrderBy(x => x.GetType().GetProperty(columnName).GetValue(x));
            }
        }
    }
}
