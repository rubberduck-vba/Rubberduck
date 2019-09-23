using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution
{
    /// <summary>
    /// Holds a context for executing a procedure scope.
    /// </summary>
    public class ExecutionContext : IExecutionContext
    {
        public ExecutionContext() : this(false) { }

        public ExecutionContext(bool isErrorPath)
        {
            IsErrorPath = isErrorPath;
        }

        public bool IsErrorPath { get; }
    }
}
