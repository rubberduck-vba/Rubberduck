using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
            var subStmts = ParseTreeResults.OfType<QualifiedContext<VBAParser.ArgListContext>>()
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                .Select(context => (VBAParser.SubStmtContext)context.Context.Parent)
                .ToList();

            var subStmtsNotImplementingInterfaces = subStmts
                .Where(c =>
            {
                var declaration =
                    UserDeclarations.SingleOrDefault(d => d.Context == c);

                if (UserDeclarations.FindInterfaceMembers().Contains(declaration))
                {
                    return false;
                }

                var interfaceImplementation = UserDeclarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
                if (interfaceImplementation == null)
                {
                    return true;
                }

                var interfaceMember = UserDeclarations.FindInterfaceMember(interfaceImplementation);

                return interfaceMember == null;
            });

            var subStmtsNotImplementingEvents = subStmts
                .Where(c =>
                {
                    var declaration = UserDeclarations.SingleOrDefault(d => d.Context == c);

                    if (declaration == null) { return false; }  // rather be safe than sorry

                    return UserDeclarations.Where(item => item.IsWithEvents)
                            .All(withEvents => UserDeclarations.FindEventProcedures(withEvents) == null) &&
                            !State.AllDeclarations.FindBuiltInEventHandlers().Contains(declaration);
                });

            return ParseTreeResults
                .Where(result => result.Context.Parent is VBAParser.SubStmtContext &&
                                  subStmtsNotImplementingInterfaces.Contains(result.Context.Parent) &&
                                  subStmtsNotImplementingEvents.Contains(result.Context.Parent)
                        && !IsInspectionDisabled(result.ModuleName.Component, result.Context.Start.Line))
                .Select(result => new ProcedureCanBeWrittenAsFunctionInspectionResult(this, State, result,
                    new QualifiedContext<VBAParser.SubStmtContext>(result.ModuleName, result.Context.Parent as VBAParser.SubStmtContext)));
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
