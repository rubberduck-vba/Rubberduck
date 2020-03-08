using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ImplementInterface;

namespace Rubberduck.Refactorings.AddInterfaceImplementations
{
    public class AddInterfaceImplementationsRefactoringAction : CodeOnlyRefactoringActionBase<AddInterfaceImplementationsModel>
    {
        private const string MemberBody = "    Err.Raise 5 'TODO implement interface member";

        public AddInterfaceImplementationsRefactoringAction(IRewritingManager rewritingManager) 
            : base(rewritingManager)
        {}

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
                return mbed.AsCodeBlock(accessibility: Tokens.Private, newIdentifier: $"{interfaceName}_{member.IdentifierName}", content: MemberBody );
            }

            if (member.DeclarationType.Equals(DeclarationType.Variable))
            {
                var propertyGet = member.FieldToPropertyBlock(DeclarationType.PropertyGet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs");
                var members = new List<string> { propertyGet };

                if (member.AsTypeName.Equals(Tokens.Variant) || !member.IsObject)
                {
                    var propertyLet = member.FieldToPropertyBlock(DeclarationType.PropertyLet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs");
                    members.Add(propertyLet);
                }

                if (member.AsTypeName.Equals(Tokens.Variant) || member.IsObject)
                {
                    var propertySet = member.FieldToPropertyBlock(DeclarationType.PropertySet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs");
                    members.Add(propertySet);
                }

                return string.Join(Environment.NewLine, members);
            }

            return string.Empty;
        }
    }
}