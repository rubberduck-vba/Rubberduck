using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

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
        public ObsoleteCommentSyntaxInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new ObsoleteCommentSyntaxListener();
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.ObsoleteCommentSyntaxInspection;
        }

        public class ObsoleteCommentSyntaxListener : InspectionListenerBase
        {
            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                SaveContext(context);
            }
        }
    }
}
