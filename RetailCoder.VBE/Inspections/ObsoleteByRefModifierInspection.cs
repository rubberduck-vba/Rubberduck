using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteByRefModifierInspection : InspectionBase
    {
        public ObsoleteByRefModifierInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteByRefModifierInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteByRefModifierInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var builtinEventHandlers = State.DeclarationFinder.FindEventHandlers();

            var allIssues = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(item =>
                    item.IsByRef && !item.IsImplicitByRef
                    && !IsIgnoringInspectionResultFor(item, AnnotationName)
                    && !builtinEventHandlers.Contains(item.ParentDeclaration)).ToArray();

            var interfaceMembers = UserDeclarations.FindInterfaceMembers();
            var interfaceIssuesDictionary = new Dictionary<ParameterDeclaration, IEnumerable<ParameterDeclaration>>();

            foreach (var member in interfaceMembers)
            {
                var interfaceIssue = allIssues.FirstOrDefault(issue => Equals(issue.ParentDeclaration, member));

                if (interfaceIssue != null)
                {
                    var implementationMembers = UserDeclarations.FindInterfaceImplementationMembers(member);
                    var implementationIssues = allIssues.Where(issue => implementationMembers.Contains(issue.ParentDeclaration));

                    interfaceIssuesDictionary.Add(interfaceIssue, implementationIssues);
                }
            }

            var interfaceIssues = interfaceIssuesDictionary.Select(x => x.Key).Concat(interfaceIssuesDictionary.SelectMany(x => x.Value));
            var nonInterfaceIssues = allIssues.Except(interfaceIssues);

            return interfaceIssuesDictionary
                .Select(item => new ObsoleteByRefModifierInspectionResult(this, item.Key, item.Value))
                .Concat(nonInterfaceIssues.Select(issue => new ObsoleteByRefModifierInspectionResult(this, issue)));
        }
    }
}
