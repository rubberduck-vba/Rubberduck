using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitByRefParameterInspection : InspectionBase
    {
        public ImplicitByRefParameterInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitByRefParameterInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitByRefParameterInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var interfaceMembersScope = UserDeclarations.FindInterfaceImplementationMembers().Select(m => m.Scope);
            var builtinEventHandlers = Declarations.FindBuiltInEventHandlers();

            var issues = (from item in UserDeclarations
                where item.DeclarationType == DeclarationType.Parameter
                    // ParamArray parameters do not allow an explicit "ByRef" parameter mechanism.               
                    && !((ParameterDeclaration)item).IsParamArray
                    && !interfaceMembersScope.Contains(item.ParentScope)
                    && !builtinEventHandlers.Contains(item.ParentDeclaration)
                let arg = item.Context as VBAParser.ArgContext
                where arg != null && arg.BYREF() == null && arg.BYVAL() == null
                select new QualifiedContext<VBAParser.ArgContext>(item.QualifiedName, arg))
                .Select(issue => new ImplicitByRefParameterInspectionResult(this, issue.Context.unrestrictedIdentifier().GetText(), issue));

            return issues;
        }
    }
}
