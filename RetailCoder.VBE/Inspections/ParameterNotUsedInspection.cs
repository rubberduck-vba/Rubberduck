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
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class ParameterNotUsedInspection : InspectionBase
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public ParameterNotUsedInspection(VBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(state)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _wrapperFactory = new CodePaneWrapperFactory();
        }

        public override string Meta { get { return InspectionsUI.ParameterNotUsedInspectionName; }}
        public override string Description { get { return InspectionsUI.ParameterNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var declarations = Declarations.ToList();

            var interfaceMemberScopes = declarations.FindInterfaceMembers().Select(m => m.Scope).ToList();
            var interfaceImplementationMemberScopes = declarations.FindInterfaceImplementationMembers().Select(m => m.Scope).ToList();

            var builtInHandlers = declarations.FindBuiltInEventHandlers();

            var parameters = declarations.Where(parameter => parameter.DeclarationType == DeclarationType.Parameter
                && !(parameter.Context.Parent.Parent is VBAParser.EventStmtContext)
                && !(parameter.Context.Parent.Parent is VBAParser.DeclareStmtContext));

            var unused = parameters.Where(parameter => !parameter.References.Any()).ToList();
            var editor = new ActiveCodePaneEditor(_vbe, _wrapperFactory);
            var quickFixRefactoring =
                new RemoveParametersRefactoring(
                    new RemoveParametersPresenterFactory(editor, 
                        new RemoveParametersDialog(), State, _messageBox), editor);

            var issues = from issue in unused.Where(parameter =>
                !IsInterfaceMemberParameter(parameter, interfaceMemberScopes)
                && !builtInHandlers.Contains(parameter.ParentDeclaration))
                let isInterfaceImplementationMember = IsInterfaceMemberImplementationParameter(issue, interfaceImplementationMemberScopes)
                select new ParameterNotUsedInspectionResult(this, string.Format(Description, issue.IdentifierName),
                        ((dynamic) issue.Context).ambiguousIdentifier(), issue.QualifiedName,
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