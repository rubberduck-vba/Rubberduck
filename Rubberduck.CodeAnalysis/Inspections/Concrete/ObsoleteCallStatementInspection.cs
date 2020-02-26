using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.ComManagement;

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
        private readonly IProjectsProvider _projectsProvider;

        public ObsoleteCallStatementInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider)
        {
            Listener = new ObsoleteCallStatementListener();
            _projectsProvider = projectsProvider;
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.ObsoleteCallStatementInspection;
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
        {
            //FIXME At least use a parse tree here instead of the COM API.
            string lines;
            var component = _projectsProvider.Component(context.ModuleName);
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

            return !stringStrippedLines.Contains(":");
        }

        public class ObsoleteCallStatementListener : InspectionListenerBase
        {
            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
