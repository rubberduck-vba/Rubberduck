using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class Parameter
    {
        public string Name { get; set; }
        public int Index { get; set; }

        public Parameter(string name, int index)
        {
            Name = name;
            Index = index;
        }
    }
}
