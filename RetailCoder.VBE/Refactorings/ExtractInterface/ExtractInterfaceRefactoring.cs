using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<ExtractInterfacePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(IRefactoringPresenterFactory<ExtractInterfacePresenter> factory,
            IActiveCodePaneEditor editor)
        {
            _factory = factory;
            _editor = editor;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();

            if (_model == null) { return; }

            AddInterface();
        }

        public void Refactor(QualifiedSelection target)
        {
            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            _editor.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void AddInterface()
        {
            var interfaceComponent = _model.TargetDeclaration.Project.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            interfaceComponent.Name = _model.InterfaceName;

            _editor.InsertLines(1, GetInterface());

            var module = _model.TargetDeclaration.QualifiedSelection.QualifiedName.Component.CodeModule;

            AddItems(module);
            module.InsertLines(module.CountOfDeclarationLines + 1, "Implements " + _model.InterfaceName);
        }

        private void AddItems(CodeModule module)
        {
            _model.Members.Reverse();

            foreach (var member in _model.Members.Where(m => m.IsSelected))
            {
                module.InsertLines(module.CountOfDeclarationLines + 1, GetInterfaceMember(member));
            }
        }

        private string GetInterfaceMember(InterfaceMember member)
        {
            switch (member.MemberType)
            {
                case "Sub":
                    return SubStmt(member);

                case "Function":
                    return FunctionStmt(member);

                case "Property":
                    switch (member.PropertyType)
                    {
                        case "Get":
                            return PropertyGetStmt(member);

                        case "Let":
                            return PropertyLetStmt(member);

                        case "Set":
                            return PropertySetStmt(member);
                    }
                    break;
            }

            return string.Empty;
        }

        private string SubStmt(InterfaceMember member)
        {
            var memberSignature = "Public Sub " + _model.InterfaceName + "_" + member.Member.IdentifierName + "(" +
                                  string.Join(", ", member.MemberParams) + ")";

            var memberBody = "    " + member.Member.IdentifierName + " " +
                             string.Join(", ", member.MemberParams.Select(p => p.ParamName));

            var memberCloseStatement = "End Sub" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, memberBody, memberCloseStatement);
        }

        private string FunctionStmt(InterfaceMember member)
        {
            var memberSignature = "Public Function " + _model.InterfaceName + "_" + member.Member.IdentifierName + "(" +
                                  string.Join(", ", member.MemberParams) + ")" + " As " + member.Type;

            var memberBody = "    " + _model.InterfaceName + "_" + member.Member.IdentifierName + " = " +
                             member.Member.IdentifierName + "(" +
                             string.Join(", ", member.MemberParams.Select(p => p.ParamName)) + ")";

            var memberCloseStatement = "End Function" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, memberBody, memberCloseStatement);
        }

        private string PropertyGetStmt(InterfaceMember member)
        {
            var memberSignature = "Public Property Get " + _model.InterfaceName + "_" + member.Member.IdentifierName +
                                  " As " + member.Type;

            var memberBody = _model.PrimitiveTypes.Contains(member.Type) ||
                                member.Member.DeclarationType == DeclarationType.UserDefinedType
                ? "    "
                : "    Set ";

            memberBody += _model.InterfaceName + "_" + member.Member.IdentifierName + " = " +
                          member.Member.IdentifierName;

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, memberBody, memberCloseStatement);
        }

        private string PropertyLetStmt(InterfaceMember member)
        {
            var memberSignature = "Public Property Let " + _model.InterfaceName + "_" + member.Member.IdentifierName +
                                  "(" + string.Join(", ", member.MemberParams) + ")";

            var memberBody = "    " + _model.InterfaceName + "_" + member.Member.IdentifierName + " = " +
                             member.Member.IdentifierName;

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, memberBody, memberCloseStatement);
        }

        private string PropertySetStmt(InterfaceMember member)
        {
            var memberSignature = "Public Property Set " + _model.InterfaceName + "_" + member.Member.IdentifierName +
                                  "(" + string.Join(", ", member.MemberParams) + ")";

            var memberBody = "    Set " + _model.InterfaceName + "_" + member.Member.IdentifierName + " = " +
                             member.Member.IdentifierName;

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, memberBody, memberCloseStatement);
        }

        private string GetInterface()
        {
            return "Option Explicit" + Environment.NewLine + string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected));
        }
    }
}