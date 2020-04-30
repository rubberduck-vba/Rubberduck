using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;


namespace Rubberduck.Refactorings.AddInterfaceImplementations
{
    public class AddInterfaceImplementationsRefactoringAction : CodeOnlyRefactoringActionBase<AddInterfaceImplementationsModel>
    {
        private readonly string _memberBody;
        private readonly ICodeBuilder _codeBuilder;

        public AddInterfaceImplementationsRefactoringAction(IRewritingManager rewritingManager, ICodeBuilder codeBuilder) 
            : base(rewritingManager)
        {
            _codeBuilder = codeBuilder;
            _memberBody = $"    {Tokens.Err}.Raise 5 {Resources.Refactorings.Refactorings.ImplementInterface_TODO}";
        }

        public override void Refactor(AddInterfaceImplementationsModel model, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetModule);
            AddItems(model.Members, rewriter, model.InterfaceName);
        }

        private void AddItems(IEnumerable<Declaration> missingMembers, IModuleRewriter rewriter, string interfaceName)
        {
            var missingMembersText = missingMembers
                .Aggregate(string.Empty, (current, member) => current + Environment.NewLine + GetInterfaceMember(member, interfaceName));

            rewriter.InsertAfter(rewriter.TokenStream.Size, Environment.NewLine + missingMembersText);
        }

        private string GetInterfaceMember(Declaration member, string interfaceName)
        {
            if (member is ModuleBodyElementDeclaration mbed)
            {
                return _codeBuilder.BuildMemberBlockFromPrototype(mbed, accessibility: Tokens.Private, newIdentifier: $"{interfaceName}_{member.IdentifierName}", content: _memberBody);
            }

            if (member is VariableDeclaration variable)
            {
                if (!_codeBuilder.TryBuildPropertyGetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertyGet, Tokens.Private, _memberBody))
                {
                    throw new InvalidOperationException();
                }

                var members = new List<string> { propertyGet };

                if (variable.AsTypeName.Equals(Tokens.Variant) || !variable.IsObject)
                {
                    if (!_codeBuilder.TryBuildPropertyLetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertyLet, Tokens.Private, _memberBody))
                    {
                        throw new InvalidOperationException();
                    }
                    members.Add(propertyLet);
                }

                if (variable.AsTypeName.Equals(Tokens.Variant) || variable.IsObject)
                {
                    if (!_codeBuilder.TryBuildPropertySetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertySet, Tokens.Private, _memberBody))
                    {
                        throw new InvalidOperationException();
                    }
                    members.Add(propertySet);
                }

                return string.Join($"{Environment.NewLine}{Environment.NewLine}", members);
            }

            return string.Empty;
        }
    }
}