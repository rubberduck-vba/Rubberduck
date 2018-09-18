using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
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
        private ClassModuleDeclaration _targetInterface;
        private ClassModuleDeclaration _targetClass;

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
            QualifiedSelection? qualifiedSelection;
            using (var activePane = _vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    _messageBox.NotifyWarn(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption);
                    return;
                }

                qualifiedSelection = activePane.GetQualifiedSelection();
                if (!qualifiedSelection.HasValue)
                {
                    _messageBox.NotifyWarn(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption);
                    return;
                }
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
            _targetInterface = _state.DeclarationFinder.FindInterface(selection);

            _targetClass = _declarations.SingleOrDefault(d =>
                        ImplementingModuleTypes.Contains(d.DeclarationType) &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName)) as ClassModuleDeclaration;

            if (_targetClass == null || _targetInterface == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.ImplementInterface_InvalidSelectionMessage, RubberduckUI.ImplementInterface_Caption);
                return;
            }

            var oldSelection = _vbe.GetActiveSelection();

            ImplementMissingMembers(_state.GetRewriter(_targetClass));

            if (oldSelection.HasValue)
            {
                var component = _state.ProjectsProvider.Component(oldSelection.Value.QualifiedName);
                using (var module = component.CodeModule)
                {
                    using (var pane = module.CodePane)
                    {
                        pane.Selection = oldSelection.Value.Selection;
                    }
                }
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
            var implemented = _targetClass.Members
                .Where(decl => decl is ModuleBodyElementDeclaration member && ReferenceEquals(member.InterfaceImplemented, _targetInterface))
                .Cast<ModuleBodyElementDeclaration>()
                .Select(member => member.InterfaceMemberImplemented).ToList();

            var interfaceMembers = _targetInterface.Members.OrderBy(member => member.Selection.StartLine)
                .ThenBy(member => member.Selection.StartColumn);

            var nonImplementedMembers = interfaceMembers.Where(member => !implemented.Contains(member));

            AddItems(nonImplementedMembers, rewriter, _targetInterface.IdentifierName);
        }

        private void AddItems(IEnumerable<Declaration> missingMembers, IModuleRewriter rewriter, string interfaceName)
        {
            var missingMembersText = missingMembers.Aggregate(string.Empty,
                (current, member) => current + Environment.NewLine + GetInterfaceMember(member, interfaceName));
            
            rewriter.InsertAfter(rewriter.TokenStream.Size, Environment.NewLine + missingMembersText);

            rewriter.Rewrite();
        }

        private string GetInterfaceMember(Declaration member, string interfaceName)
        {
            var template = string.Join(Environment.NewLine, Tokens.Private + " {0}{1} {2}{3}", MemberBody, Tokens.End + " {0}", string.Empty);
            var signature = $"{interfaceName}_{member.IdentifierName}({string.Join(", ", GetParameters(member))})";
            var asType = $" {Tokens.As} {member.AsTypeName}";

            switch (member.DeclarationType)
            {
                case DeclarationType.Procedure:
                    return string.Format(template, Tokens.Sub, string.Empty, signature, string.Empty);
                case DeclarationType.Function:
                    return string.Format(template, Tokens.Function, string.Empty, signature, asType);
                case DeclarationType.PropertyGet:
                    return string.Format(template, Tokens.Property, $" {Tokens.Get}", signature, asType);
                case DeclarationType.PropertyLet:
                    return string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty);
                case DeclarationType.PropertySet:
                    return string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty);
                case DeclarationType.Variable:
                    var members = new List<string>
                    {
                        string.Format(template, Tokens.Property, $" {Tokens.Get}", $"{interfaceName}_{member.IdentifierName}()", asType)
                    };

                    if (member.AsTypeName.Equals(Tokens.Variant) || !member.IsObject)
                    {
                        members.Add(string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty));
                    }

                    if (member.AsTypeName.Equals(Tokens.Variant) || member.IsObject)
                    {
                        members.Add(string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty));
                    }

                    return string.Join(Environment.NewLine, members);
            }

            return string.Empty;
        }

        private IEnumerable<Parameter> GetParameters(Declaration member)
        {
            if (member.DeclarationType == DeclarationType.Variable)
            {
                return new List<Parameter>
                {
                    new Parameter
                    {
                        Accessibility = Tokens.ByVal,
                        Name = "rhs",
                        AsTypeName = member.AsTypeName
                    }
                };
            }

            return ((ModuleBodyElementDeclaration) member).Parameters.Select(p => new Parameter
            {
                Accessibility = ((VBAParser.ArgContext) p.Context).BYVAL() != null
                    ? Tokens.ByVal
                    : Tokens.ByRef,
                Name = p.IdentifierName,
                AsTypeName = p.AsTypeName
            });
        }
    }
}
