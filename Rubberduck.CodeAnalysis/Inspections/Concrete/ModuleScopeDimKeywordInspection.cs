using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about module-level declarations made using the 'Dim' keyword.
    /// </summary>
    /// <why>
    /// Private module variables should be declared using the 'Private' keyword. While 'Dim' is also legal, it should preferably be 
    /// restricted to declarations of procedure-scoped local variables, for consistency, since public module variables are declared with the 'Public' keyword.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Dim foo As Long
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Private foo As Long
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class ModuleScopeDimKeywordInspection : ParseTreeInspectionBase
    {
        public ModuleScopeDimKeywordInspection(RubberduckParserState state) 
            : base(state) { }

        public override IInspectionListener Listener { get; } = new ModuleScopedDimListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .SelectMany(result => result.Context.GetDescendents<VBAParser.VariableSubStmtContext>()
                        .Select(r => new QualifiedContext<ParserRuleContext>(result.ModuleName, r)))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       string.Format(InspectionResults.ModuleScopeDimKeywordInspection, ((VBAParser.VariableSubStmtContext)result.Context).identifier().GetText()),
                                                       result));
        }

        public class ModuleScopedDimListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitVariableStmt([NotNull] VBAParser.VariableStmtContext context)
            {
                if (context.DIM() != null && context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out _))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}