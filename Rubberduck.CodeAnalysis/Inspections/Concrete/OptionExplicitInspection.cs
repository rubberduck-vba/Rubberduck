using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
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
        public OptionExplicitInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new MissingOptionExplicitListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this,
                    string.Format(InspectionResults.OptionExplicitInspection, context.ModuleName.ComponentName),
                    context));
        }

        public class MissingOptionExplicitListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly IDictionary<string, QualifiedContext<ParserRuleContext>> _contexts = new Dictionary<string,QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts.Values.ToList();

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitModuleBody(VBAParser.ModuleBodyContext context)
            {
                if (context.ChildCount == 0 && _contexts.ContainsKey(CurrentModuleName.Name))
                {
                    _contexts.Remove(CurrentModuleName.Name);
                }
            }

            public override void ExitModuleDeclarations([NotNull] VBAParser.ModuleDeclarationsContext context)
            {
                var hasOptionExplicit = false;
                foreach (var element in context.moduleDeclarationsElement())
                {
                    if (element.moduleOption() is VBAParser.OptionExplicitStmtContext)
                    {
                        hasOptionExplicit = true;
                    }
                }

                if (!hasOptionExplicit)
                {
                    _contexts.Add(CurrentModuleName.Name, new QualifiedContext<ParserRuleContext>(CurrentModuleName, (ParserRuleContext)context.Parent));
                }
            }
        }
    }
}
