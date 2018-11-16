using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class SymbolTable<TKey, TValue>
    {
        private readonly Dictionary<TKey, TValue> _table;

        public SymbolTable()
        {
            _table = new Dictionary<TKey, TValue>();
        }

        public void AddOrUpdate(TKey name, TValue value)
        {
            _table[name] = value;
        }

        public bool HasSymbol(TKey name)
        {
            return _table.ContainsKey(name);            
        }

        public TValue Get(TKey name)
        {
            if (_table.ContainsKey(name))
            {
                return _table[name];
            }
            throw new InvalidOperationException(name + " not found in symbol table.");
        }
    }
}
