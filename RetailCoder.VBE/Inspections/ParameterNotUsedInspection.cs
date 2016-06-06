using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class ParameterNotUsedInspection : InspectionBase
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;

        public ParameterNotUsedInspection(VBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(state)
        {
            _vbe = vbe;
            _messageBox = messageBox;
        }

        public override string Meta { get { return InspectionsUI.ParameterNotUsedInspectionName; }}
        public override string Description { get { return InspectionsUI.ParameterNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var interfaceMemberScopes = declarations.FindInterfaceMembers().Select(m => m.Scope).ToList();
            var interfaceImplementationMemberScopes = declarations.FindInterfaceImplementationMembers().Select(m => m.Scope).ToList();

            var builtInHandlers = declarations.FindBuiltInEventHandlers();

            var parameters = declarations.Where(parameter => parameter.DeclarationType == DeclarationType.Parameter
                && parameter.ParentDeclaration.DeclarationType != DeclarationType.Event
                && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryFunction
                && parameter.ParentDeclaration.DeclarationType != DeclarationType.LibraryProcedure);

            var unused = parameters.Where(parameter => !parameter.References.Any()).ToList();
            var quickFixRefactoring =
                new RemoveParametersRefactoring(_vbe, new RemoveParametersPresenterFactory(_vbe, new RemoveParametersDialog(), State, _messageBox));

            var issues = from issue in unused.Where(parameter =>
                !IsInterfaceMemberParameter(parameter, interfaceMemberScopes)
                && !builtInHandlers.Contains(parameter.ParentDeclaration))
                let isInterfaceImplementationMember = IsInterfaceMemberImplementationParameter(issue, interfaceImplementationMemberScopes)
                select new ParameterNotUsedInspectionResult(this, issue,
                        ((dynamic) issue.Context).unrestrictedIdentifier(), issue.QualifiedName,
                        isInterfaceImplementationMember, quickFixRefactoring, State);

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
