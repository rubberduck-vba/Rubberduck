using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ProcedureCanBeWrittenAsFunctionInspection : InspectionBase
    {
        public ProcedureCanBeWrittenAsFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ProcedureShouldBeFunctionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ProcedureCanBeFunctionInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var subStmts = State.ArgListsWithOneByRefParam
                .Where(context => context.Context.Parent is VBAParser.SubStmtContext)
                .Select(context => (VBAParser.SubStmtContext)context.Context.Parent)
                .ToList();

            var subStmtsNotImplementingInterfaces = subStmts
                .Where(c =>
            {
                var declaration =
                    UserDeclarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Procedure &&
                                                          d.IdentifierName == c.ambiguousIdentifier().GetText() &&
                                                          d.Context.GetSelection().Equals(c.GetSelection()));

                var interfaceImplementation = UserDeclarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));

                if (interfaceImplementation == null) { return true; }

                var interfaceMember = UserDeclarations.FindInterfaceMember(interfaceImplementation);
                return interfaceMember == null;
            });

            var subStmtsNotImplementingEvents = subStmts
                .Where(c =>
                {
                    var declaration = UserDeclarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Procedure &&
                                                              d.IdentifierName == c.ambiguousIdentifier().GetText() &&
                                                              d.Context.GetSelection().Equals(c.GetSelection()));

                    if (declaration == null) { return false; }  // rather be safe than sorry

                    return UserDeclarations.Where(item => item.IsWithEvents)
                            .All(withEvents => UserDeclarations.FindEventProcedures(withEvents) == null);
                });

            return State.ArgListsWithOneByRefParam
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

        private bool IsInterfaceImplementation(Declaration target)
        {
            var interfaceImplementation = State.AllUserDeclarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(target));

            if (interfaceImplementation == null) { return false; }

            var interfaceMember = State.AllUserDeclarations.FindInterfaceMember(interfaceImplementation);
            return interfaceMember != null;
        }
    }
}
