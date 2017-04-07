using System.Collections.Generic;
using System.Linq;
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

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : InspectionBase, IParseTreeInspection
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (_parseTreeResults == null)
            {
                Logger.Debug("Aborting GetInspectionResults because ParseTree results were not passed");
                return new InspectionResultBase[] { };
            }

            var userDeclarations = UserDeclarations.ToList();
            var builtinHandlers = State.DeclarationFinder.FindEventHandlers().ToList();

            var contextLookup = userDeclarations.Where(decl => decl.Context != null).ToDictionary(decl => decl.Context);

            var ignored = new HashSet<Declaration>( State.DeclarationFinder.FindAllInterfaceMembers()
                .Concat(State.DeclarationFinder.FindAllInterfaceImplementingMembers())
                .Concat(builtinHandlers)
                .Concat(userDeclarations.Where(item => item.IsWithEvents)));

            return _parseTreeResults.Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                                   .Select(context => contextLookup[(VBAParser.SubStmtContext)context.Context.Parent])
                                   .Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName) &&
                                                  !ignored.Contains(decl) &&
                                                  userDeclarations.Where(item => item.IsWithEvents)
                                                                  .All(withEvents => userDeclarations.FindEventProcedures(withEvents) == null) &&
                                                                  !builtinHandlers.Contains(decl))
                                   .Select(result => new ProcedureCanBeWrittenAsFunctionInspectionResult(
                                                         this,
                                                         State,
                                                         result,
                                                         new QualifiedContext<VBAParser.SubStmtContext>(result.QualifiedName, (VBAParser.SubStmtContext)result.Context))
                                   );                   
        }

        public class SingleByRefParamArgListListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ArgListContext> _contexts = new List<VBAParser.ArgListContext>();
            public IEnumerable<VBAParser.ArgListContext> Contexts => _contexts;

            public override void ExitArgList(VBAParser.ArgListContext context)
            {
                var args = context.arg();
                if (args != null && args.All(a => a.PARAMARRAY() == null && a.LPAREN() == null) && args.Count(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)) == 1)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
