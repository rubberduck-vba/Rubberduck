using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        private readonly List<Declaration> _declarations;
        private Declaration _targetInterface;
        private Declaration _targetClass;

        private const string MemberBody = "    Err.Raise 5 'TODO implement interface member";

        public ImplementInterfaceRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
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

        private static readonly IReadOnlyList<DeclarationType> ImplementingModuleTypes = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.Document, 
        };

        public void Refactor(QualifiedSelection selection)
        {
            _targetInterface = _declarations.FindInterface(selection);

            _targetClass = _declarations.SingleOrDefault(d =>
                        ImplementingModuleTypes.Contains(d.DeclarationType) &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName));

            if (_targetClass == null || _targetInterface == null)
            {
                _messageBox.Show(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            QualifiedSelection? oldSelection = null;
            if (_vbe.ActiveCodePane != null)
            {
                oldSelection = _vbe.ActiveCodePane.CodeModule.GetQualifiedSelection();
            }

            ImplementMissingMembers(_state.GetRewriter(_targetClass));

            if (oldSelection.HasValue)
            {
                var module = oldSelection.Value.QualifiedName.Component.CodeModule;
                var pane = module.CodePane;
                pane.Selection = oldSelection.Value.Selection;
            }

            _state.OnParseRequested(this);
        }

        public void Refactor(Declaration target)
        {
            throw new NotSupportedException();
        }

        internal void Refactor(List<Declaration> members, IModuleRewriter rewriter, string interfaceName)
        {
            AddItems(members, rewriter, interfaceName);
        }

        private void ImplementMissingMembers(IModuleRewriter rewriter)
        {
            var interfaceMembers = GetInterfaceMembers();
            var implementedMembers = GetImplementedMembers();
            var nonImplementedMembers = GetNonImplementedMembers(interfaceMembers, implementedMembers);

            AddItems(nonImplementedMembers, rewriter, _targetInterface.IdentifierName);
        }

        private void AddItems(List<Declaration> missingMembers, IModuleRewriter rewriter, string interfaceName)
        {
            var missingMembersText = missingMembers.Aggregate(string.Empty,
                (current, member) => current + Environment.NewLine + GetInterfaceMember(member, interfaceName));
            
            rewriter.InsertAfter(rewriter.TokenStream.Size, Environment.NewLine + missingMembersText);

            rewriter.Rewrite();
        }

        private string GetInterfaceMember(Declaration member, string interfaceName)
        {
            switch (member.DeclarationType)
            {
                case DeclarationType.Procedure:
                    return SubStmt(member, interfaceName);

                case DeclarationType.Function:
                    return FunctionStmt(member, interfaceName);

                case DeclarationType.PropertyGet:
                    return PropertyGetStmt(member, interfaceName);

                case DeclarationType.PropertyLet:
                    return PropertyLetStmt(member, interfaceName);

                case DeclarationType.PropertySet:
                    return PropertySetStmt(member, interfaceName);
            }

            return string.Empty;
        }

        private string SubStmt(Declaration member, string interfaceName)
        {
            var memberParams = GetParameters(member);

            var memberSignature = $"Private Sub {interfaceName}_{member.IdentifierName}({string.Join(", ", memberParams)})";
            var memberCloseStatement = "End Sub" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string FunctionStmt(Declaration member, string interfaceName)
        {
            var memberParams = GetParameters(member);

            var memberSignature = $"Private Function {interfaceName}_{member.IdentifierName}({string.Join(", ", memberParams)}) As {member.AsTypeName}";
            var memberCloseStatement = "End Function" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyGetStmt(Declaration member, string interfaceName)
        {
            var memberParams = GetParameters(member);

            var memberSignature = $"Private Property Get {interfaceName}_{member.IdentifierName}({string.Join(", ", memberParams)}) As {member.AsTypeName}";
            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertyLetStmt(Declaration member, string interfaceName)
        {
            var memberParams = GetParameters(member);

            var memberSignature = $"Private Property Let {interfaceName}_{member.IdentifierName}({string.Join(", ", memberParams)})";
            var memberCloseStatement = "End Property" + Environment.NewLine;

            return string.Join(Environment.NewLine, memberSignature, MemberBody, memberCloseStatement);
        }

        private string PropertySetStmt(Declaration member, string interfaceName)
        {
            var memberParams = GetParameters(member);

            var memberSignature = $"Private Property Set {interfaceName}_{member.IdentifierName}({string.Join(", ", memberParams)})";
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
    }
}
