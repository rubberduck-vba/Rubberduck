using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

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

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var interfaceMembers = UserDeclarations.FindInterfaceImplementationMembers();

            var issues = (from item in UserDeclarations
                where !item.IsInspectionDisabled(AnnotationName)
                    && item.DeclarationType == DeclarationType.Parameter
                    && !interfaceMembers.Select(m => m.Scope).Contains(item.ParentScope)
                let arg = item.Context as VBAParser.ArgContext
                where arg != null && arg.BYREF() == null && arg.BYVAL() == null
                select new QualifiedContext<VBAParser.ArgContext>(item.QualifiedName, arg))
                .Select(issue => new ImplicitByRefParameterInspectionResult(this, string.Format(Description, issue.Context.ambiguousIdentifier().GetText()), issue));

 
            return issues;
        }
    }
}