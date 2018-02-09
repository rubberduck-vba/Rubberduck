using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class ContextExtents<T> where T : IComparable<T>
    {
        private T _minValue;
        private T _maxValue;
        bool _hasValues;

        public ContextExtents()
        {
            _minValue = default;
            _maxValue = default;
            _hasValues = false;
        }

        public void MinMax(T minVal, T maxVal)
        {
            _minValue = minVal;
            _maxValue = maxVal;
            _hasValues = true;
        }

        public T Min => _minValue;
        public T Max => _maxValue;
        public bool HasValues => _hasValues;
    }
}
