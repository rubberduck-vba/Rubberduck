using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : InspectionBase, IParseTreeInspection
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public IInspectionListener Listener { get; } =
            new SingleByRefParamArgListListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var userDeclarations = UserDeclarations.ToList();
            var builtinHandlers = State.DeclarationFinder.FindEventHandlers().ToList();

            var contextLookup = userDeclarations.Where(decl => decl.Context != null).ToDictionary(decl => decl.Context);

            var ignored = new HashSet<Declaration>( State.DeclarationFinder.FindAllInterfaceMembers()
                .Concat(State.DeclarationFinder.FindAllInterfaceImplementingMembers())
                .Concat(builtinHandlers)
                .Concat(userDeclarations.Where(item => item.IsWithEvents)));

            return Listener.Contexts.Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                                   .Select(context => contextLookup[(VBAParser.SubStmtContext)context.Context.Parent])
                                   .Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName) &&
                                                  !ignored.Contains(decl) &&
                                                  userDeclarations.Where(item => item.IsWithEvents)
                                                                  .All(withEvents => userDeclarations.FindEventProcedures(withEvents) == null) &&
                                                                  !builtinHandlers.Contains(decl))
                                   .Select(result => new DeclarationInspectionResult(this,
                                                             string.Format(InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionResultFormat, result.IdentifierName),
                                                             result));                   
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
