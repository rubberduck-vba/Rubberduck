using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;
using NLog;

namespace Rubberduck.Inspections
{
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : InspectionBase, IParseTreeInspection
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ProcedureCanBeWrittenAsFunctionInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                _logger.Debug("Aborting GetInspectionResults because ParseTree results were not passed");
                return new InspectionResultBase[] { };
            }
            var subStmts = ParseTreeResults.ArgListsWithOneByRefParam
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                .Select(context => (VBAParser.SubStmtContext)context.Context.Parent)
                .ToList();

            var subStmtsNotImplementingInterfaces = subStmts
                .Where(c =>
            {
                var declaration =
                    UserDeclarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Procedure &&
                                                          d.IdentifierName == c.subroutineName().GetText() &&
                                                          d.Context.GetSelection().Equals(c.GetSelection()));

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
                    var declaration = UserDeclarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Procedure &&
                                                              d.IdentifierName == c.subroutineName().GetText() &&
                                                              d.Context.GetSelection().Equals(c.GetSelection()));

                    if (declaration == null) { return false; }  // rather be safe than sorry

                    return UserDeclarations.Where(item => item.IsWithEvents)
                            .All(withEvents => UserDeclarations.FindEventProcedures(withEvents) == null) &&
                            !State.AllDeclarations.FindBuiltInEventHandlers().Contains(declaration);
                });

            return ParseTreeResults.ArgListsWithOneByRefParam
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext &&
                                  subStmtsNotImplementingInterfaces.Contains(context.Context.Parent) &&
                                  subStmtsNotImplementingEvents.Contains(context.Context.Parent))
                .Select(context => new ProcedureShouldBeFunctionInspectionResult(this,
                    State,
                    new QualifiedContext<VBAParser.ArgListContext>(context.ModuleName,
                        context.Context as VBAParser.ArgListContext),
                    new QualifiedContext<VBAParser.SubStmtContext>(context.ModuleName,
                        context.Context.Parent as VBAParser.SubStmtContext)));
        }

        public class ArgListWithOneByRefParamListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.ArgListContext> _contexts = new List<VBAParser.ArgListContext>();
            public IEnumerable<VBAParser.ArgListContext> Contexts { get { return _contexts; } }

            public override void ExitArgList(VBAParser.ArgListContext context)
            {
                if (context.arg() != null && context.arg().Count(a => a.BYREF() != null || (a.BYREF() == null && a.BYVAL() == null)) == 1)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
