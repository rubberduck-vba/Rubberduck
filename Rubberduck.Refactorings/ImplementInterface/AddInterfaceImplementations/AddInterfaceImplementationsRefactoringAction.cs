using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.Refactorings.AddInterfaceImplementations
{
    public class AddInterfaceImplementationsRefactoringAction : CodeOnlyRefactoringActionBase<AddInterfaceImplementationsModel>
    {
        private readonly ICodeBuilder _codeBuilder;

        public AddInterfaceImplementationsRefactoringAction(IRewritingManager rewritingManager, ICodeBuilder codeBuilder) 
            : base(rewritingManager)
        {
            _codeBuilder = codeBuilder;
        }

        public override void Refactor(AddInterfaceImplementationsModel model, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetModule);

            var missingMembersText = model.Members
                .Aggregate(string.Empty, (current, member) => current + NewLines.DOUBLE_SPACE + GetInterfaceMember(member, model.InterfaceName, $"{model.GetMemberImplementation(member)}"));

            rewriter.InsertAfter(rewriter.TokenStream.Size, missingMembersText);
        }

        private string GetInterfaceMember(Declaration member, string interfaceName, string memberBody)
        {
            var implementingMemberName = $"{interfaceName}_{member.IdentifierName}";

            if (member is ModuleBodyElementDeclaration mbed)
            {
                return _codeBuilder.BuildMemberBlockFromPrototype(mbed, accessibility: Accessibility.Private, newIdentifier: $"{interfaceName}_{member.IdentifierName}", content: memberBody);
            }

            if (member is VariableDeclaration variable)
            {
                if (!_codeBuilder.TryBuildPropertyGetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertyGet, Accessibility.Private, memberBody))
                {
                    throw new InvalidOperationException();
                }

                var members = new List<string> { propertyGet };

                if (variable.AsTypeName.Equals(Tokens.Variant) || !variable.IsObject)
                {
                    if (!_codeBuilder.TryBuildPropertyLetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertyLet, Accessibility.Private, memberBody))
                    {
                        throw new InvalidOperationException();
                    }
                    members.Add(propertyLet);
                }

                if (variable.AsTypeName.Equals(Tokens.Variant) || variable.IsObject)
                {
                    if (!_codeBuilder.TryBuildPropertySetCodeBlock(variable, $"{interfaceName}_{variable.IdentifierName}", out var propertySet, Accessibility.Private, memberBody))
                    {
                        throw new InvalidOperationException();
                    }
                    members.Add(propertySet);
                }

                return string.Join($"{NewLines.DOUBLE_SPACE}", members);
            }

            return string.Empty;
        }
    }
}