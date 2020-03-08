using System.Collections.Generic;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags modules that omit Option Explicit.
    /// </summary>
    /// <why>
    /// This option makes variable declarations mandatory. Without it, a typo gets compiled as a new on-the-spot Variant/Empty variable with a new name. 
    /// Omitting this option amounts to refusing the little help the VBE can provide with compile-time validation.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    ///
    /// 
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class OptionExplicitInspection : ParseTreeInspectionBase<VBAParser.ModuleContext>
    {
        public OptionExplicitInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new MissingOptionExplicitListener();
        }

        protected  override IInspectionListener<VBAParser.ModuleContext> ContextListener { get; }

        protected override bool IsResultContext(QualifiedContext<VBAParser.ModuleContext> context, DeclarationFinder finder)
        {
            var moduleBody = context.Context.moduleBody();
            return moduleBody != null && moduleBody.ChildCount != 0;
        }

        protected override string ResultDescription(QualifiedContext<VBAParser.ModuleContext> context)
        {
            var moduleName = context.ModuleName.ComponentName;
            return string.Format(
                InspectionResults.OptionExplicitInspection,
                moduleName);
        }

        private class MissingOptionExplicitListener : InspectionListenerBase<VBAParser.ModuleContext>
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
                    SaveContext((VBAParser.ModuleContext)context.Parent);
                }
            }
        }
    }
}
