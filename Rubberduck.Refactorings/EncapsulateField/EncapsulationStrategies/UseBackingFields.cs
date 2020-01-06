using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UseBackingFields : EncapsulateFieldStrategyBase
    {
        //private IEnumerable<IEncapsulateFieldCandidate> _convertedFields;
        public UseBackingFields(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter)
            : base(declarationFinderProvider, model, indenter)
        {
            //_convertedFields = model.SelectedFieldCandidates; //.Cast<IUsingBackingField>().ToList();
            model.AssignCandidateValidations(EncapsulateFieldStrategy.UseBackingFields);
            //foreach (var candidate in _convertedFields)
            //{
            //    if (candidate is IUserDefinedTypeCandidate)
            //    {
            //        candidate.NameValidator = model.ValidatorProvider.NameOnlyValidator(Validators.UserDefinedType);
            //    }
            //    else if (candidate is IUserDefinedTypeMemberCandidate)
            //    {
            //        candidate.NameValidator = candidate.Declaration.IsArray
            //            ? model.ValidatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMemberArray)
            //            : model.ValidatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMember);
            //    }
            //    else
            //    {
            //        candidate.NameValidator = model.ValidatorProvider.NameOnlyValidator(Validators.Default);
            //    }
            //    candidate.ConflictFinder = model.ValidatorProvider.ConflictDetector(EncapsulateFieldStrategy.UseBackingFields, declarationFinderProvider);
            //}
        }

        protected override void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in SelectedFields)
            {
                if (field.Declaration.HasPrivateAccessibility() && field.FieldIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList() && !field.Declaration.HasPrivateAccessibility())
                {
                    refactorRewriteSession.Remove(field.Declaration, rewriter);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.FieldIdentifier);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        protected override void ModifyReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            foreach (var field in SelectedFields)
            {
                field.LoadFieldReferenceContextReplacements();
            }

            RewriteReferences(model, refactorRewriteSession);
        }

        protected override void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = SelectedFields
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.FieldIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddContentBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }
    }
}
