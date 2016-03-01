using System.Collections.Generic;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class SymbolTable
    {
        private readonly Dictionary<string, object> _table;

        public SymbolTable()
        {
            _table = new Dictionary<string, object>();
        }

        public void Add(string name, object value)
        {
            _table[name] = value;
        }

        public bool HasSymbol(string name)
        {
            return _table.ContainsKey(name);            
        }

        public object Get(string name)
        {
            if (_table.ContainsKey(name))
            {
                return _table[name];
            }
            return null;
        }
    }
}
