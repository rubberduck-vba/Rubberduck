using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceRefactoring : RefactoringBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private const string MemberBody = "    Err.Raise 5 'TODO implement interface member";

        public ImplementInterfaceRefactoring(
            IDeclarationFinderProvider declarationFinderProvider, 
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider)
        :base(rewritingManager, selectionProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        private static readonly IReadOnlyList<DeclarationType> ImplementingModuleTypes = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.Document
        };

        public override void Refactor(QualifiedSelection target)
        {
            var targetInterface = _declarationFinderProvider.DeclarationFinder.FindInterface(target);

            if (targetInterface == null)
            {
                throw new NoImplementsStatementSelectedException(target);
            }

            var targetModule = _declarationFinderProvider.DeclarationFinder
                .ModuleDeclaration(target.QualifiedName);
            
            if (!ImplementingModuleTypes.Contains(targetModule.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(targetModule);
            }

            var targetClass = targetModule as ClassModuleDeclaration;

            if (targetClass == null)
            {
                //This really should never happen. If it happens the declaration type enum value
                //and the type of the declaration are inconsistent.
                throw new InvalidTargetDeclarationException(targetModule);
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(targetClass.QualifiedModuleName);
            ImplementMissingMembers(targetInterface, targetClass, rewriter);
            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            throw new NotSupportedException();
        }

        public override void Refactor(Declaration target)
        {
            throw new NotSupportedException();
        }

        internal void Refactor(List<Declaration> members, IModuleRewriter rewriter, string interfaceName)
        {
            AddItems(members, rewriter, interfaceName);
        }

        private void ImplementMissingMembers(ModuleDeclaration targetInterface, ModuleDeclaration targetClass, IModuleRewriter rewriter)
        {
            var implemented = targetClass.Members
                .Where(decl => decl is ModuleBodyElementDeclaration member && ReferenceEquals(member.InterfaceImplemented, targetInterface))
                .Cast<ModuleBodyElementDeclaration>()
                .Select(member => member.InterfaceMemberImplemented).ToList();

            var interfaceMembers = targetInterface.Members.OrderBy(member => member.Selection.StartLine)
                .ThenBy(member => member.Selection.StartColumn);

            var nonImplementedMembers = interfaceMembers.Where(member => !implemented.Contains(member));

            AddItems(nonImplementedMembers, rewriter, targetInterface.IdentifierName);
        }

        private void AddItems(IEnumerable<Declaration> missingMembers, IModuleRewriter rewriter, string interfaceName)
        {
            var missingMembersText = missingMembers.Aggregate(string.Empty,
                (current, member) => current + Environment.NewLine + GetInterfaceMember(member, interfaceName));
            
            rewriter.InsertAfter(rewriter.TokenStream.Size, Environment.NewLine + missingMembersText);
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
