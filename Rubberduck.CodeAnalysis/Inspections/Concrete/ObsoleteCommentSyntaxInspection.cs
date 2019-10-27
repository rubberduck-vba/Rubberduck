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
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Rem' comments.
    /// </summary>
    /// <why>
    /// Modern VB comments use a single quote character (') to denote the beginning of a comment: the legacy 'Rem' syntax is obsolete.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// Rem this comment is using an obsolete legacy syntax
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// ' this comment is using the modern comment syntax
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteCommentSyntaxInspection : ParseTreeInspectionBase
    {
        public ObsoleteCommentSyntaxInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteCommentSyntaxListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this, InspectionResults.ObsoleteCommentSyntaxInspection, context));
        }

        public class ObsoleteCommentSyntaxListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}
