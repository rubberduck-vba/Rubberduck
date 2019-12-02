using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField.Strategies
{
    public interface IEncapsulateFieldStrategy
    {
        IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
        IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
        Dictionary<string, IEncapsulateFieldCandidate> UdtMemberTargetIDToParentMap { get; set; }
        Dictionary<string, IEncapsulateFieldCandidate> FlattenedTargetIDToCandidateMapping { get; }
    }

    public abstract class EncapsulateFieldStrategiesBase : IEncapsulateFieldStrategy
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected EncapsulationCandidateFactory _candidateFactory;
        private Dictionary<string, IEncapsulateFieldCandidate> _udtMemberTargetIDToParentMap { get; } = new Dictionary<string, IEncapsulateFieldCandidate>();
        private IEncapsulateFieldNamesValidator _validator;


        public EncapsulateFieldStrategiesBase(QualifiedModuleName qmn, IIndenter indenter, IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldNamesValidator validator)
        {
            TargetQMN = qmn;
            Indenter = indenter;
            _declarationFinderProvider = declarationFinderProvider;
            _validator = validator;
            _candidateFactory = new EncapsulationCandidateFactory(declarationFinderProvider, _validator);

            EncapsulationCandidateFields = _declarationFinderProvider.DeclarationFinder
                .Members(qmn)
                .Where(v => v.IsMemberVariable() && !v.IsWithEvents);

            var candidates = _candidateFactory.CreateEncapsulationCandidates(EncapsulationCandidateFields);
            foreach (var candidate in candidates)
            {
                HeirarchicalCandidates.Add(candidate.TargetID, candidate);
            }

            FlattenedTargetIDToCandidateMapping = Flatten(HeirarchicalCandidates);
            foreach (var element in FlattenedTargetIDToCandidateMapping)
            {
                if (element.Value.IsUDTMember)
                {
                    var udtMember = element.Value as EncapsulatedUserDefinedTypeMember;
                    UdtMemberTargetIDToParentMap.Add(element.Key, udtMember.Parent);
                }
            }
        }

        protected QualifiedModuleName TargetQMN {private set; get;}

        protected IIndenter Indenter { private set; get; }

        public IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            return RefactorRewrite(model, rewriteSession, true);
        }

        public IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            return RefactorRewrite(model, rewriteSession, false);
        }

        public Dictionary<string, IEncapsulateFieldCandidate> FlattenedTargetIDToCandidateMapping { get; } = new Dictionary<string, IEncapsulateFieldCandidate>();

        protected virtual IEnumerable<Declaration> EncapsulationCandidateFields { set; get; }

        protected abstract void ModifyEncapsulatedVariable(IEncapsulateFieldCandidate target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession); //, EncapsulateFieldNewContent newContent)

        protected abstract EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields);

        protected abstract IList<string> PropertiesContent(IEnumerable<IEncapsulateFieldCandidate> flaggedEncapsulationFields);

        protected Dictionary<string, IEncapsulateFieldCandidate> HeirarchicalCandidates { set; get; } = new Dictionary<string, IEncapsulateFieldCandidate>();

        private IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(encFld => encFld.Declaration.IsVariable());

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                ModifyEncapsulatedVariable(nonUdtMemberField, attributes, rewriteSession);
                RenameReferences(nonUdtMemberField, attributes.PropertyName ?? nonUdtMemberField.Declaration.IdentifierName, rewriteSession);
            }

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);
            RewriterRemoveWorkAround.RemoveDeclarationsFromVariableLists(rewriter);

            InsertNewContent(model.CodeSectionStartIndex, model, rewriteSession, asPreview);

            return rewriteSession;
        }

        public Dictionary<string, IEncapsulateFieldCandidate> UdtMemberTargetIDToParentMap { get; set; } = new Dictionary<string, IEncapsulateFieldCandidate>();

        private void InsertNewContent(int? codeSectionStartIndex, EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool includePreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            var newContent = new EncapsulateFieldNewContent();
            newContent = LoadNewDeclarationsContent(newContent, model.FlaggedEncapsulationFields);

            if (includePreviewMessage)
            {
                var postScript = "'<===== No Changes below this line =====>";
                newContent = LoadNewPropertiesContent(newContent, model.FlaggedEncapsulationFields, postScript);
            }
            else
            {
                newContent = LoadNewPropertiesContent(newContent, model.FlaggedEncapsulationFields);
            }

            rewriter.InsertNewContent(codeSectionStartIndex, newContent);

        }

        protected Dictionary<string, IEncapsulateFieldCandidate> Flatten(Dictionary<string, IEncapsulateFieldCandidate> heirarchicalCandidates)
        {
            var candidates = new Dictionary<string, IEncapsulateFieldCandidate>();
            foreach (var keyValue in heirarchicalCandidates)
            {
                candidates.Add(keyValue.Key, keyValue.Value);
                if (keyValue.Value.Declaration.IsUserDefinedTypeField())
                {
                    if (keyValue.Value is EncapsulatedUserDefinedTypeField udt)
                    {
                        foreach (var member in udt.Members)
                        {
                            candidates.Add(member.TargetID, member);
                            _udtMemberTargetIDToParentMap.Add(member.TargetID, keyValue.Value);
                        }
                    }
                }
            }
            return candidates;
        }

        private EncapsulateFieldNewContent LoadNewPropertiesContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields, string postScript = null)
        {
            if (!FlaggedEncapsulationFields.Any()) { return newContent; }

            var theContent = string.Join($"{Environment.NewLine}{Environment.NewLine}", PropertiesContent(FlaggedEncapsulationFields));
            newContent.AddCodeBlock(theContent);
            if (postScript?.Length > 0)
            {
                newContent.AddCodeBlock($"{postScript}{Environment.NewLine}{Environment.NewLine}");
            }
            return newContent;
        }

        private void RenameReferences(IEncapsulateFieldCandidate efd, string propertyName, IRewriteSession rewriteSession)
        {
            foreach (var reference in efd.Declaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, propertyName ?? efd.Declaration.IdentifierName);
            }
        }
    }
}
