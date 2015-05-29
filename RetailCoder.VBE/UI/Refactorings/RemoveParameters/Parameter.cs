using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Remove Parameters refactoring.
    /// </summary>
    public class Parameter
    {
        public Declaration Declaration { get; private set; }
        public string Name { get; private set; }
        public int Index { get; private set; }
        public bool IsRemoved { get; set; }

        public Parameter(Declaration declaration, int index, bool isRemoved = false)
        {
            Declaration = declaration;
            Name = declaration.Context.GetText();
            Index = index;
            IsRemoved = isRemoved;
        }
    }
}
