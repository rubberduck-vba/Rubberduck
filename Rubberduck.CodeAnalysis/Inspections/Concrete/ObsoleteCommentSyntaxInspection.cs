using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Rem' comments.
    /// </summary>
    /// <why>
    /// Modern VB comments use a single quote character (') to denote the beginning of a comment: the legacy 'Rem' syntax is obsolete.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// Rem this comment is using an obsolete legacy syntax
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// ' this comment is using the modern comment syntax
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObsoleteCommentSyntaxInspection : ParseTreeInspectionBase<VBAParser.RemCommentContext>
    {
        public ObsoleteCommentSyntaxInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ObsoleteCommentSyntaxListener();
        }

        protected override IInspectionListener<VBAParser.RemCommentContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.RemCommentContext> context)
        {
            return InspectionResults.ObsoleteCommentSyntaxInspection;
        }

        private class ObsoleteCommentSyntaxListener : InspectionListenerBase<VBAParser.RemCommentContext>
        {
            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                SaveContext(context);
            }
        }
    }
}
