using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    public interface IAssignmentNode : IExecutableNode
    {
        IdentifierReference Target { get; set; }

    }
}
