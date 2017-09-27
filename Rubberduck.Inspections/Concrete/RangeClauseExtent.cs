using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public interface IRangeClause
    {
        bool IsSingleVal { get; }
        bool IsRange { get; }
        bool UsesIsClause { get; }
        string ValueAsString { get; }
        string ValueMinAsString { get; }
        string ValueMaxAsString { get; }
        string TypeName { get; }
        string CompareSymbol { get; }
    }

    public class RangeClauseExtent<T> : IComparable, IRangeClause where T : System.IComparable
    {
        public RangeClauseExtent( T extent, string typeName, string compareSymbol)
        {
            _extent = extent;
            _typeName = typeName;
            _compareSymbol = compareSymbol;
        }

        private T _extent;
        private string _typeName;
        private string _compareSymbol;

        public bool IsSingleVal => true;

        public bool IsRange => false;

        public bool UsesIsClause => true;

        public string ValueAsString => _extent.ToString();

        public string TypeName => _typeName;

        public string CompareSymbol => _compareSymbol;

        public string ValueMinAsString => throw new NotImplementedException();

        public string ValueMaxAsString => throw new NotImplementedException();

        public int CompareTo(object x)
        {
            return 1;
        }


    }
}
