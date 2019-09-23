using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution
{
    public partial class ExecutionEngine
    {
        /// <summary>
        /// Traverses the executable nodes of the specified member.
        /// </summary>
        /// <param name="member">The member to "execute".</param>
        /// <returns>The execution context for each code path in the specified member.</returns>
        public IEnumerable<IExecutionContext> Execute(ModuleBodyElementDeclaration member)
        {
            return Enumerable.Empty<IExecutionContext>();
        }
    }
}
