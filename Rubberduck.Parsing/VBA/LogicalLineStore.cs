using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.VBA
{
    public class LogicalLineStore
    {
        private readonly List<int> _logicalLineEnds;

        public LogicalLineStore(IEnumerable<int> logicalLineEnds)
        {
            _logicalLineEnds = logicalLineEnds.ToList();
            if (!IsSorted(_logicalLineEnds))
            {
                _logicalLineEnds.Sort();
            }
        }

        private static bool IsSorted(List<int> list)
        {
            for (var index = 0; index < list.Count - 1; index++)
            {
                if (list[index + 1] < list[index])
                {
                    return false;
                }
            }

            return true;
        }

        private int LastPhysicalLine() => _logicalLineEnds[_logicalLineEnds.Count - 1];

        public int? LogicalLineNumber(int physicalLineNumber)
        {
            if (physicalLineNumber < 1 || physicalLineNumber > LastPhysicalLine())
            {
                return null;
            }

            var searchResult = _logicalLineEnds.BinarySearch(physicalLineNumber);
            if (searchResult >= 0)
            {
                //VBA line numbers are 1-based.
                return searchResult + 1;
            }

            //If the item is not in the sorted list, BinarySearch returns the bitwise complement of the index with the next larger item,
            //in case there is one, and the binary complement of the length or the list otherwise.
            var nextLargerIndex = ~searchResult;

            //VBA line numbers are 1-based.
            return nextLargerIndex + 1;
        }

        public int? PhysicalEndLineNumber(int logicalLineNumber)
        {
            if (logicalLineNumber < 1 || logicalLineNumber > NumberOfLogicalLines())
            {
                return null;
            }

            //VBA line numbers are 1-based.
            return _logicalLineEnds[logicalLineNumber - 1];
        }

        public int? PhysicalStartLineNumber(int logicalLineNumber)
        {
            if (logicalLineNumber < 1 || logicalLineNumber > NumberOfLogicalLines())
            {
                return null;
            }

            //VBA line numbers are 1-based. So the prior logical line has an index two smaller than the logical line number.
            return logicalLineNumber == 1 
                ? 1
                : _logicalLineEnds[logicalLineNumber - 2] + 1;
        }

        public int NumberOfLogicalLines()
        {
            return _logicalLineEnds.Count;
        }

        public int? StartOfContainingLogicalLine(int physicalLineNumber)
        {
            var logicalLine = LogicalLineNumber(physicalLineNumber);
            return logicalLine.HasValue
                ? PhysicalStartLineNumber(logicalLine.Value)
                : null;
        }
        public int? EndOfContainingLogicalLine(int physicalLineNumber)
        {
            var logicalLine = LogicalLineNumber(physicalLineNumber);
            return logicalLine.HasValue
                ? PhysicalEndLineNumber(logicalLine.Value)
                : null;
        }
    }
}