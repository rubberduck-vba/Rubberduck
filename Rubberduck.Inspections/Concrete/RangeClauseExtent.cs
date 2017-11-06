using System;

namespace Rubberduck.Inspections.Concrete
{
    public interface IRangeClause
    {
        bool IsSingleVal { get; }
        bool IsRange { get; }
        bool UsesIsClause { get; }
        bool IsRangeExtent { get; }
        string ValueAsString { get; }
        string ValueMinAsString { get; }
        string ValueMaxAsString { get; }
        string CompareSymbol { get; }
    }

    public class RangeClauseExtent<T> : IRangeClause
    {
        private T _extent;
        private string _compareSymbol;

        public RangeClauseExtent(T extent, string compareSymbol)
        {
            _extent = extent;
            _compareSymbol = compareSymbol;
        }

        public bool IsSingleVal => true;
        public bool IsRange => false;
        public bool UsesIsClause => true;
        public bool IsRangeExtent => true;
        public string ValueAsString => _extent.ToString();
        public string CompareSymbol => _compareSymbol;
        public string ValueMinAsString => ValueAsString;
        public string ValueMaxAsString => ValueAsString;
    }
}
