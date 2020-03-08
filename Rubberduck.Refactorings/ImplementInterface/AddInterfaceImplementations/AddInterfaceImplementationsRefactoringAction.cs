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
                var members = new List<string>
                {
                    member.FieldToPropertyBlock(DeclarationType.PropertyGet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs")
                };

                if (member.AsTypeName.Equals(Tokens.Variant) || !member.IsObject)
                {
                    members.Add(member.FieldToPropertyBlock(DeclarationType.PropertyLet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs"));
                    //members.Add(string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty));
                }

                if (member.AsTypeName.Equals(Tokens.Variant) || member.IsObject)
                {
                    members.Add(member.FieldToPropertyBlock(DeclarationType.PropertySet, $"{interfaceName}_{member.IdentifierName}", Tokens.Private, MemberBody, "rhs"));
                    //members.Add(string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty));
                }


                //var template = string.Join(Environment.NewLine, Tokens.Private + " {0}{1} {2}{3}", MemberBody, Tokens.End + " {0}", string.Empty);
                //var signature = $"{interfaceName}_{member.IdentifierName}({string.Join(", ", GetParameters(member))})";
                //var asType = $" {Tokens.As} {member.AsTypeName}";
                //var members = new List<string>
                //    {
                //        string.Format(template, Tokens.Property, $" {Tokens.Get}", $"{interfaceName}_{member.IdentifierName}()", asType)
                //    };

                //if (member.AsTypeName.Equals(Tokens.Variant) || !member.IsObject)
                //{
                //    members.Add(string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty));
                //}

                //if (member.AsTypeName.Equals(Tokens.Variant) || member.IsObject)
                //{
                //    members.Add(string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty));
                //}

                return string.Join(Environment.NewLine, members);
            }

            return string.Empty;

            //var template = string.Join(Environment.NewLine, Tokens.Private + " {0}{1} {2}{3}", MemberBody, Tokens.End + " {0}", string.Empty);
            //var signature = $"{interfaceName}_{member.IdentifierName}({string.Join(", ", GetParameters(member))})";
            //var asType = $" {Tokens.As} {member.AsTypeName}";

            //switch (member.DeclarationType)
            //{
            //    case DeclarationType.Procedure:
            //        return string.Format(template, Tokens.Sub, string.Empty, signature, string.Empty);
            //    case DeclarationType.Function:
            //        return string.Format(template, Tokens.Function, string.Empty, signature, asType);
            //    case DeclarationType.PropertyGet:
            //        return string.Format(template, Tokens.Property, $" {Tokens.Get}", signature, asType);
            //    case DeclarationType.PropertyLet:
            //        return string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty);
            //    case DeclarationType.PropertySet:
            //        return string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty);
            //    case DeclarationType.Variable:
            //        var members = new List<string>
            //        {
            //            string.Format(template, Tokens.Property, $" {Tokens.Get}", $"{interfaceName}_{member.IdentifierName}()", asType)
            //        };

            //        if (member.AsTypeName.Equals(Tokens.Variant) || !member.IsObject)
            //        {
            //            members.Add(string.Format(template, Tokens.Property, $" {Tokens.Let}", signature, string.Empty));
            //        }

            //        if (member.AsTypeName.Equals(Tokens.Variant) || member.IsObject)
            //        {
            //            members.Add(string.Format(template, Tokens.Property, $" {Tokens.Set}", signature, string.Empty));
            //        }

            //        return string.Join(Environment.NewLine, members);
            //}

            //return string.Empty;
        }

        //private IEnumerable<Parameter> GetParameters(Declaration member)
        //{
        //    if (member.DeclarationType == DeclarationType.Variable)
        //    {
        //        return new List<Parameter>
        //        {
        //            new Parameter
        //            {
        //                Accessibility = Tokens.ByVal,
        //                Name = "rhs",
        //                AsTypeName = member.AsTypeName
        //            }
        //        };
        //    }

        //    if (member is ModuleBodyElementDeclaration method)
        //    {
        //        return method.Parameters
        //            .OrderBy(parameter => parameter.Selection)
        //            .Select(parameter => new Parameter(parameter));
        //    }

        //    return Enumerable.Empty<Parameter>();
        //}
    }
}