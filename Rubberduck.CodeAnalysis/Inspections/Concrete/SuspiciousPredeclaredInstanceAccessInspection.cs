using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// This inspection warns about references to the default instance of a class, inside that class.
    /// </summary>
    /// <why>
    /// While a stateful default instance might be intentional, it is a common source of bugs and should be avoided.
    /// Use the Me qualifier to explicitly refer to the current instance and eliminate any ambiguity.
    /// </why>
    /// <example hasResult="true">
    /// <module name="UserForm1" type="UserForm Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private ClickCount As Long
    /// 
    /// Private Sub CommandButton1_Click()
    ///     ClickCount = ClickCount + 1
    ///     UserForm1.TextBox1.Text = ClickCount
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="UserForm1" type="UserForm Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private ClickCount As Long
    /// 
    /// Private Sub CommandButton1_Click()
    ///     ClickCount = ClickCount + 1
    ///     Me.TextBox1.Text = ClickCount
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class SuspiciousPredeclaredInstanceAccessInspection : IdentifierReferenceInspectionBase
    {
        public SuspiciousPredeclaredInstanceAccessInspection(IDeclarationFinderProvider declarationFinderProvider) 
            : base(declarationFinderProvider)
        {
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return 
                reference.Declaration is ClassModuleDeclaration module && 
                module.HasPredeclaredId &&
                reference.ParentScoping.ParentDeclaration.Equals(module) &&
                reference.Context.TryGetAncestor<VBAParser.LExpressionContext>(out var expression) &&
                reference.IdentifierName != Tokens.Me && expression.GetText() == reference.IdentifierName;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            reference.Context.TryGetAncestor<VBAParser.LExpressionContext>(out var expression);
            return string.Format(InspectionResults.SuspiciousPredeclaredInstanceAccessInspection, reference.IdentifierName, expression.GetText());
        }
    }
}
