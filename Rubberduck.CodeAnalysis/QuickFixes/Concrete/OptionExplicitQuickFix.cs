using System;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds 'Option Explicit' to the top of code modules.
    /// </summary>
    /// <inspections>
    /// <inspection name="OptionExplicitInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class OptionExplicitQuickFix : QuickFixBase
    {
        public OptionExplicitQuickFix()
            : base(typeof(OptionExplicitInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(0, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.OptionExplicitQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}