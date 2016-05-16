using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
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
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        private List<Declaration> _declarations;
        private Declaration _targetInterface;
        private Declaration _targetClass;

        private const string MemberBody = "    Err.Raise 5 'TODO implement interface member";

        public ImplementInterfaceRefactoring(VBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _declarations = state.AllUserDeclarations.ToList();
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            if (_vbe.ActiveCodePane == null)
            {
                _messageBox.Show(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            var qualifiedSelection = _vbe.ActiveCodePane.GetQualifiedSelection();
            if (!qualifiedSelection.HasValue)
            {
                _messageBox.Show(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            Refactor(qualifiedSelection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            _targetInterface = _declarations.FindInterface(selection);

            _targetClass = _declarations.SingleOrDefault(d =>
                        !d.IsBuiltIn && d.DeclarationType == DeclarationType.ClassModule &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName));

            if (_targetClass == null || _targetInterface == null)
            {
                _messageBox.Show(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            ImplementMissingMembers();

            _state.OnParseRequested(this);
        }

        public void Refactor(Declaration target)
        {
            throw new NotImplementedException();
        }

        private void ImplementMissingMembers()
        {
            var interfaceMembers = GetInterfaceMembers();
            var implementedMembers = GetImplementedMembers();
            var nonImplementedMembers = GetNonImplementedMembers(interfaceMembers, implementedMembers);

            AddItems(nonImplementedMembers);
        }

        private void AddItems(List<Declaration> members)
        {
            var module = _targetClass.QualifiedSelection.QualifiedName.Component.CodeModule;

            var missingMembersText = members.Aggregate(string.Empty, (current, member) => current + Environment.NewLine + GetInterfaceMember(member));

            module.InsertLines(module.CountOfDeclarationLines + 2, missingMembersText);
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

            var memberSignature = "Private Sub " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Sub" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string FunctionStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Private Function " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")" + " As " + member.AsTypeName;

            var memberCloseStatement = "End Function" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyGetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Private Property Get " + _targetInterface.IdentifierName + "_" + member.IdentifierName + "(" +
                                  string.Join(", ", memberParams) + ")" + " As " + member.AsTypeName;

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyLetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Private Property Let " + _targetInterface.IdentifierName + "_" + member.IdentifierName +
                                  "(" + string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertySetStmt(Declaration member)
        {
            var memberParams = GetParameters(member);

            var memberSignature = "Private Property Set " + _targetInterface.IdentifierName + "_" + member.IdentifierName +
                                  "(" + string.Join(", ", memberParams) + ")";

            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private List<Parameter> GetParameters(Declaration member)
        {
            var parameters = _declarations.Where(item => item.DeclarationType == DeclarationType.Parameter &&
                              item.ParentScopeDeclaration == member)
                           .OrderBy(o => o.Selection.StartLine)
                           .ThenBy(t => t.Selection.StartColumn)
                           .Select(p => new Parameter
                           {
                               Accessibility = ((VBAParser.ArgContext)p.Context).BYVAL() != null
                                            ? Tokens.ByVal 
                                            : Tokens.ByRef,

                               Name = p.IdentifierName,
                               AsTypeName = p.AsTypeName
                           })
                           .ToList();

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
                                .Where(item => item.ProjectId == _targetInterface.ProjectId
                                        && item.ComponentName == _targetClass.IdentifierName
                                        && item.IdentifierName.StartsWith(_targetInterface.ComponentName + "_")
                                        && !item.Equals(_targetClass))
                                .OrderBy(d => d.Selection.StartLine)
                                .ThenBy(d => d.Selection.StartColumn);
        }

        private List<Declaration> GetNonImplementedMembers(IEnumerable<Declaration> interfaceMembers, IEnumerable<Declaration> implementedMembers)
        {
            return interfaceMembers.Where(d => !implementedMembers.Select(s => s.IdentifierName)
                                        .Contains(_targetInterface.ComponentName + "_" + d.IdentifierName))
                                    .OrderBy(o => o.Selection.StartLine)
                                    .ThenBy(t => t.Selection.StartColumn)
                                    .ToList();
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