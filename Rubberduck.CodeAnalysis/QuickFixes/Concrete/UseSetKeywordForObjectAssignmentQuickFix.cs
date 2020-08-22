using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces a 'Set' keyword for what appears to be a suspicious or malformed object reference assignment.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObjectVariableNotSetInspection" />
    /// <inspection name="SuspiciousLetAssignmentInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim c As VBA.Collection
    ///     c = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim c As VBA.Collection
    ///     Set c = New VBA.Collection
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class UseSetKeywordForObjectAssignmentQuickFix : QuickFixBase
    {
        public UseSetKeywordForObjectAssignmentQuickFix()
            : base(typeof(ObjectVariableNotSetInspection), typeof(SuspiciousLetAssignmentInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var letStmt = result.Context.GetAncestor<VBAParser.LetStmtContext>();
            var letToken = letStmt.LET();
            var setToken = Tokens.Set;
            if (letToken != null)
            {
                rewriter.Replace(letToken, setToken);
            }
            else
            {
                setToken += " ";
                rewriter.InsertBefore(letStmt.Start.TokenIndex, setToken);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SetObjectVariableQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}