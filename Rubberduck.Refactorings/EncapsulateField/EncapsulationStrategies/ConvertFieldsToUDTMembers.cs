using Antlr4.Runtime;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ConvertFieldsToUDTMembers : EncapsulateFieldStrategyBase
    {
        public ConvertFieldsToUDTMembers(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName qmn, IIndenter indenter)
            : base(declarationFinderProvider, qmn, indenter) { }

        protected override void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in model.SelectedFieldCandidates)
            {
                refactorRewriteSession.Remove(field.Declaration, rewriter);
            }

            if (model.StateUDTField.IsExistingDeclaration)
            {
                model.StateUDTField.AddMembers(model.SelectedFieldCandidates);

                rewriter.Replace(model.StateUDTField.AsTypeDeclaration, model.StateUDTField.TypeDeclarationBlock(_indenter));
            }
        }

        protected override void ModifyReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            foreach (var field in model.SelectedFieldCandidates)
            {
                field.LoadFieldReferenceContextReplacements(model.StateUDTField.FieldIdentifier);
            }

            RewriteReferences(model, refactorRewriteSession);
        }

        protected override void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            if (model.StateUDTField.IsExistingDeclaration) { return; }

            model.StateUDTField.AddMembers(model.SelectedFieldCandidates);

            AddContentBlock(NewContentTypes.TypeDeclarationBlock, model.StateUDTField.TypeDeclarationBlock(_indenter));

            AddContentBlock(NewContentTypes.DeclarationBlock, model.StateUDTField.FieldDeclarationBlock);
            return;
        }
    }
}
