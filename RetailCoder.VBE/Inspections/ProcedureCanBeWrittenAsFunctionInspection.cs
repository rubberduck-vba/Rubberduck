using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;

namespace Rubberduck.Inspections
{
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : InspectionBase, IParseTreeInspection
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private IEnumerable<QualifiedContext> _results;

        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public IEnumerable<QualifiedContext<VBAParser.ArgListContext>> ParseTreeResults { get { return _results.OfType<QualifiedContext<VBAParser.ArgListContext>>(); } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _results = results;
        }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
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

            return ParseTreeResults.Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                                   .Select(context => contextLookup[(VBAParser.SubStmtContext)context.Context.Parent])
                                   .Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName) &&
                                                  !ignored.Contains(decl) &&
                                                  userDeclarations.Where(item => item.IsWithEvents)
                                                                  .All(withEvents => userDeclarations.FindEventProcedures(withEvents) == null) &&
                                                                  !builtinHandlers.Contains(decl))
                                   .Select(result => new ProcedureCanBeWrittenAsFunctionInspectionResult(
                                                         this, 
                                                         State, 
                                                         new QualifiedContext<VBAParser.ArgListContext>(result.QualifiedName,result.Context.GetChild<VBAParser.ArgListContext>(0)),
                                                         new QualifiedContext<VBAParser.SubStmtContext>(result.QualifiedName, (VBAParser.SubStmtContext)result.Context))
                                   );                   
        }

        public class SingleByRefParamArgListListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ArgListContext> _contexts = new List<VBAParser.ArgListContext>();
            public IEnumerable<VBAParser.ArgListContext> Contexts { get { return _contexts; } }

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
