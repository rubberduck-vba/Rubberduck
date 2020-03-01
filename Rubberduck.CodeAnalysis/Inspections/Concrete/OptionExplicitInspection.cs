using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags modules that omit Option Explicit.
    /// </summary>
    /// <why>
    /// This option makes variable declarations mandatory. Without it, a typo gets compiled as a new on-the-spot Variant/Empty variable with a new name. 
    /// Omitting this option amounts to refusing the little help the VBE can provide with compile-time validation.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    ///
    /// 
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class OptionExplicitInspection : ParseTreeInspectionBase
    {
        public OptionExplicitInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new MissingOptionExplicitListener();
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var moduleName = context.ModuleName.ComponentName;
            return string.Format(
                InspectionResults.OptionExplicitInspection,
                moduleName);
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
        {
            var moduleBody = (context.Context as VBAParser.ModuleContext)?.moduleBody();
            return moduleBody != null && moduleBody.ChildCount != 0;
        }

        public class MissingOptionExplicitListener : InspectionListenerBase
        {
            private readonly IDictionary<QualifiedModuleName, bool> _hasOptionExplicit = new Dictionary<QualifiedModuleName, bool>();

            public override void ClearContexts()
            {
                _hasOptionExplicit.Clear();
                base.ClearContexts();
            }

            public override void ClearContexts(QualifiedModuleName module)
            {
                _hasOptionExplicit.Remove(module);
                base.ClearContexts(module);
            }

            public override void EnterModuleDeclarations(VBAParser.ModuleDeclarationsContext context)
            {
                _hasOptionExplicit[CurrentModuleName] = false;
            }

            public override void ExitOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
            {
                _hasOptionExplicit[CurrentModuleName] = true;
            }

            public override void ExitModuleDeclarations([NotNull] VBAParser.ModuleDeclarationsContext context)
            {
                if (!_hasOptionExplicit.TryGetValue(CurrentModuleName, out var hasOptionExplicit) || !hasOptionExplicit)
                {
                    SaveContext((ParserRuleContext)context.Parent);
                }
            }
        }
    }
}
