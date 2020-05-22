using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections
{
    public interface IInspectionListener: IParseTreeListener
    {
        void ClearContexts();
        void ClearContexts(QualifiedModuleName module);
        QualifiedModuleName CurrentModuleName { get; set; }
    }

    public interface IInspectionListener<TContext> : IInspectionListener
        where TContext : ParserRuleContext
    {
        IReadOnlyList<QualifiedContext<TContext>> Contexts();
        IReadOnlyList<QualifiedContext<TContext>> Contexts(QualifiedModuleName module);
    }
}