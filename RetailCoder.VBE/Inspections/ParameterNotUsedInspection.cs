using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspection : IInspection
    {
        public ParameterNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ParameterNotUsedInspection"; } }
        public string Description { get { return RubberduckUI.ParameterNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMemberScopes = parseResult.Declarations.FindInterfaceMembers().Select(m => m.Scope).ToList();
            var interfaceImplementationMemberScopes = parseResult.Declarations.FindInterfaceImplementationMembers().Select(m => m.Scope).ToList();

            var parameters = parseResult.Declarations.Items.Where(parameter => !parameter.IsBuiltIn
                && parameter.DeclarationType == DeclarationType.Parameter
                && !(parameter.Context.Parent.Parent is VBAParser.EventStmtContext)
                && !(parameter.Context.Parent.Parent is VBAParser.DeclareStmtContext));

            var unused = parameters.Where(parameter => !parameter.References.Any()).ToList();
            var quickFixRefactoring =
                new RemoveParametersRefactoring(
                    new RemoveParametersPresenterFactory(new ActiveCodePaneEditor(parseResult.Project.VBE),
                        new RemoveParametersDialog(), parseResult));

            var issues = from issue in unused.Where(parameter => !IsInterfaceMemberParameter(parameter, interfaceMemberScopes))
                         let isInterfaceImplementationMember = IsInterfaceMemberImplementationParameter(issue, interfaceImplementationMemberScopes)
                         select new ParameterNotUsedInspectionResult(string.Format(Description, issue.IdentifierName), Severity, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName, isInterfaceImplementationMember, quickFixRefactoring, parseResult);

            return issues.ToList();
        }

        private bool IsInterfaceMemberParameter(Declaration parameter, IEnumerable<string> interfaceMemberScopes)
        {
            return interfaceMemberScopes.Contains(parameter.ParentScope);
        }

        private bool IsInterfaceMemberImplementationParameter(Declaration parameter, IEnumerable<string> interfaceMemberImplementationScopes)
        {
            return interfaceMemberImplementationScopes.Contains(parameter.ParentScope);
        }
    }
}