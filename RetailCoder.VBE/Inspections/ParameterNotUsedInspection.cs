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
    public class ParameterNotUsedInspection : IInspection
    {
        private readonly VBE _vbe;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public ParameterNotUsedInspection(VBE vbe)
        {
            _vbe = vbe; // todo: remove this dependency
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ParameterNotUsedInspection"; } }
        public string Description { get { return RubberduckUI.ParameterNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var declarations = state.AllDeclarations.ToList();

            var interfaceMemberScopes = declarations.FindInterfaceMembers().Select(m => m.Scope).ToList();
            var interfaceImplementationMemberScopes = declarations.FindInterfaceImplementationMembers().Select(m => m.Scope).ToList();

            var parameters = declarations.Where(parameter => !parameter.IsBuiltIn
                && parameter.DeclarationType == DeclarationType.Parameter
                && !(parameter.Context.Parent.Parent is VBAParser.EventStmtContext)
                && !(parameter.Context.Parent.Parent is VBAParser.DeclareStmtContext));

            var unused = parameters.Where(parameter => !parameter.References.Any()).ToList();
            var editor = new ActiveCodePaneEditor(_vbe, _wrapperFactory);
            var quickFixRefactoring =
                new RemoveParametersRefactoring(
                    new RemoveParametersPresenterFactory(editor, 
                        new RemoveParametersDialog(), state, new MessageBox()), editor);

            var issues = from issue in unused.Where(parameter => !IsInterfaceMemberParameter(parameter, interfaceMemberScopes))
                         let isInterfaceImplementationMember = IsInterfaceMemberImplementationParameter(issue, interfaceImplementationMemberScopes)
                         select new ParameterNotUsedInspectionResult(this, string.Format(Description, issue.IdentifierName), ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName, isInterfaceImplementationMember, quickFixRefactoring, state);

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