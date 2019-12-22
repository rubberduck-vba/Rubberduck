using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates explicit 'Call' statements.
    /// </summary>
    /// <why>
    /// The 'Call' keyword is obsolete and redundant, since call statements are legal and generally more consistent without it.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub Test()
    ///     Call DoSomething(42)
    /// End Sub
    ///
    /// Private Sub DoSomething(ByVal foo As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub Test()
    ///     DoSomething 42
    /// End Sub
    ///
    /// Private Sub DoSomething(ByVal foo As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteCallStatementInspection : ParseTreeInspectionBase
    {
        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteCallStatementListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();
            // do prefiltering to reduce searchspace
            var prefilteredContexts = Listener.Contexts.Where(context => !context.IsIgnoringInspectionResultFor(State.DeclarationFinder, AnnotationName));
            foreach (var context in prefilteredContexts)
            {
                string lines;
                var component = State.ProjectsProvider.Component(context.ModuleName);
                using (var module = component.CodeModule)
                {
                    lines = module.GetLines(context.Context.Start.Line,
                        context.Context.Stop.Line - context.Context.Start.Line + 1);
                }

                var stringStrippedLines = string.Join(string.Empty, lines).StripStringLiterals();

                if (stringStrippedLines.HasComment(out var commentIndex))
                {
                    stringStrippedLines = stringStrippedLines.Remove(commentIndex);
                }

                if (!stringStrippedLines.Contains(":"))
                {
                    results.Add(new QualifiedContextInspectionResult(this,
                                                     InspectionResults.ObsoleteCallStatementInspection,
                                                     context));
                }
            }

            return results;
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
