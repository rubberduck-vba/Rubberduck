using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about 'Sub' procedures that could be refactored into a 'Function'.
    /// </summary>
    /// <why>
    /// Idiomatic VB code uses 'Function' procedures to return a single value. If the procedure isn't side-effecting, consider writing is as a
    /// 'Function' rather than a 'Sub' the returns a result through a 'ByRef' parameter.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef result As Long)
    ///     ' ...
    ///     result = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Function DoSomething() As Long
    ///     ' ...
    ///     DoSomething = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : ParseTreeInspectionBase
    {
        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new SingleByRefParamArgListListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            if (!Listener.Contexts.Any())
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var userDeclarations = UserDeclarations.ToList();
            var builtinHandlers = State.DeclarationFinder.FindEventHandlers().ToList();

            var contextLookup = userDeclarations.Where(decl => decl.Context != null).ToDictionary(decl => decl.Context);

            var ignored = new HashSet<Declaration>( State.DeclarationFinder.FindAllInterfaceMembers()
                .Concat(State.DeclarationFinder.FindAllInterfaceImplementingMembers())
                .Concat(builtinHandlers)
                .Concat(userDeclarations.Where(item => item.IsWithEvents)));

            return Listener.Contexts
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext
                                    && HasArgumentReferencesWithIsAssignmentFlagged(context))
                .Select(GetSubStmtParentDeclaration)
                .Where(decl => decl != null && 
                                !ignored.Contains(decl) &&
                                userDeclarations.Where(item => item.IsWithEvents)
                                   .All(withEvents => !State.DeclarationFinder.FindHandlersForWithEventsField(withEvents).Any()) &&
                               !builtinHandlers.Contains(decl))
                .Select(result => new DeclarationInspectionResult(this,
                    string.Format(InspectionResults.ProcedureCanBeWrittenAsFunctionInspection, result.IdentifierName),
                    result));

            bool HasArgumentReferencesWithIsAssignmentFlagged(QualifiedContext<ParserRuleContext> context)
            {
                return contextLookup.TryGetValue(context.Context.GetChild<VBAParser.ArgContext>(), out Declaration decl) 
                       && decl.References.Any(rf => rf.IsAssignment);
            }

            Declaration GetSubStmtParentDeclaration(QualifiedContext<ParserRuleContext> context)
            {
                return contextLookup.TryGetValue(context.Context.Parent as VBAParser.SubStmtContext, out Declaration decl)
                    ? decl
                        : null;
            }
        }

        public class SingleByRefParamArgListListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitArgList(VBAParser.ArgListContext context)
            {
                var args = context.arg();
                if (args != null && args.All(a => a.PARAMARRAY() == null && a.LPAREN() == null) && args.Count(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)) == 1)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
