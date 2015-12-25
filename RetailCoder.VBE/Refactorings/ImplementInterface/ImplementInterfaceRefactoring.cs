using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceRefactoring : IRefactoring
    {
        private readonly List<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private Declaration _targetInterface;
        private Declaration _targetClass;
        private readonly IMessageBox _messageBox;

        const string MemberBody = "    Err.Raise 5";

        public ImplementInterfaceRefactoring(RubberduckParserState state, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _declarations = state.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _editor.GetSelection();

            if (!selection.HasValue)
            {
                _messageBox.Show("Invalid selection.", "Rubberduck - Implement Interface",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            Refactor(selection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            _targetInterface = _declarations.FindInterface(selection);

            _targetClass = _declarations.SingleOrDefault(d =>
                        !d.IsBuiltIn && d.DeclarationType == DeclarationType.Class &&
                        d.QualifiedSelection.QualifiedName.ComponentName == selection.QualifiedName.ComponentName &&
                        d.Project == selection.QualifiedName.Project);

            if (_targetClass == null || _targetInterface == null)
            {
                return;
            }

            ImplementMissingMembers();
        }

        public void Refactor(Declaration target)
        {
            throw new NotImplementedException();
        }

        private void ImplementMissingMembers()
        {
            var interfaceMembers = GetInterfaceMembers();
            var implementedMembers = GetImplementedMembers();

            var nonImplementedMembers =
                interfaceMembers.Where(
                    d =>
                        !implementedMembers.Select(s => s.IdentifierName)
                            .Contains(_targetInterface.ComponentName + "_" + d.IdentifierName)).ToList();

            AddItems(nonImplementedMembers);
        }

        private void AddItems(List<Declaration> members)
        {
            var module = _targetClass.QualifiedSelection.QualifiedName.Component.CodeModule;

            members.Reverse();

            foreach (var member in members)
            {
                module.InsertLines(module.CountOfDeclarationLines + 1, GetInterfaceMember(member));
            }
        }

        private string GetInterfaceMember(Declaration member)
        {
            switch (GetMemberType(member))
            {
                case "Sub":
                    return SubStmt(member);

                case "Function":
                    return FunctionStmt(member);

                case "Property Get":
                    return PropertyGetStmt(member);

                case "Property Let":
                    return PropertyLetStmt(member);

                case "Property Set":
                    return PropertySetStmt(member);
            }

            return string.Empty;
        }

        private string SubStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Public Sub " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Sub" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string FunctionStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Public Function " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")" + " As " + member.AsTypeName;

            var memberCloseStatement = "End Function" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyGetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Public Property Get " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")" + " As " + member.AsTypeName;

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyLetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Public Property Let " + _targetInterface.IdentifierName + "_" + member.IdentifierName +
                                  "(" + string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertySetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Public Property Set " + _targetInterface.IdentifierName + "_" + member.IdentifierName +
                                  "(" + string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private List<Parameter> GetParameters(Declaration member)
        {
            var parameters = _declarations.Where(item => item.DeclarationType == DeclarationType.Parameter &&
                              item.ParentScope == member.Scope)
                           .OrderBy(o => o.Selection.StartLine)
                           .ThenBy(t => t.Selection.StartColumn)
                           .Select(p => new Parameter
                           {
                               ParamAccessibility = ((VBAParser.ArgContext)p.Context).BYREF() == null ? Tokens.ByVal : Tokens.ByRef,
                               ParamName = p.IdentifierName,
                               ParamType = p.AsTypeName
                           })
                           .ToList();

            if (member.DeclarationType == DeclarationType.PropertyGet)
            {
                parameters.Remove(parameters.Last());
            }

            return parameters;
        }

        private IEnumerable<Declaration> GetInterfaceMembers()
        {
            return _declarations.FindInterfaceMembers()
                                .Where(d => d.ComponentName == _targetInterface.IdentifierName)
                                .OrderBy(d => d.Selection.StartLine)
                                .ThenBy(d => d.Selection.StartColumn);
        }

        private IEnumerable<Declaration> GetImplementedMembers()
        {
            return _declarations.FindInterfaceImplementationMembers()
                                .Where(item => item.Project.Equals(_targetInterface.Project)
                                        && item.ComponentName == _targetClass.IdentifierName
                                        && item.IdentifierName.StartsWith(_targetInterface.ComponentName + "_")
                                        && !item.Equals(_targetClass))
                                .OrderBy(d => d.Selection.StartLine)
                                .ThenBy(d => d.Selection.StartColumn);
        }

        private string GetMemberType(Declaration member)
        {
            var context = member.Context;

            var subStmtContext = context as VBAParser.SubStmtContext;
            if (subStmtContext != null)
            {
                return Tokens.Sub;
            }

            var functionStmtContext = context as VBAParser.FunctionStmtContext;
            if (functionStmtContext != null)
            {
                return Tokens.Function;
            }

            var propertyGetStmtContext = context as VBAParser.PropertyGetStmtContext;
            if (propertyGetStmtContext != null)
            {
                return Tokens.Property + " " + Tokens.Get;
            }

            var propertyLetStmtContext = context as VBAParser.PropertyLetStmtContext;
            if (propertyLetStmtContext != null)
            {
                return Tokens.Property + " " + Tokens.Let;
            }

            var propertySetStmtContext = context as VBAParser.PropertySetStmtContext;
            if (propertySetStmtContext != null)
            {
                return Tokens.Property + " " + Tokens.Set;
            }

            return string.Empty;
        }
    }
}