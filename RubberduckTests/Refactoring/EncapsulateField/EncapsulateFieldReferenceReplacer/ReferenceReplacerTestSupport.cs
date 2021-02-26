using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class ReferenceReplacerTestSupport
    {
        public static IDictionary<string, string> TestReferenceReplacement(bool wrapInPrivateUDT, (string, string, bool) testTargetTuple, params (string, string, ComponentType)[] moduleTuples)
        {
            var vbe = MockVbeBuilder.BuildFromModules(moduleTuples);
            return wrapInPrivateUDT
                ? ReplaceReferencesWrapInPrivateUDT(vbe.Object, testTargetTuple)
                : ReplaceReferences(vbe.Object, testTargetTuple);

        }

        private static IDictionary<string, string> ReplaceReferences(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, params (string fieldID, string fieldProperty, bool readOnly)[] fieldIDPairs)
            => ReplaceReferences(vbe, false, target, fieldIDPairs.ToList());

        private static IDictionary<string, string> ReplaceReferencesWrapInPrivateUDT(IVBE vbe, (string fieldID, string fieldProperty, bool readOnly) target, params (string fieldID, string fieldProperty, bool readOnly)[] fieldIDPairs)
            => ReplaceReferences(vbe, true, target, fieldIDPairs.ToList());

        private static IDictionary<string, string> ReplaceReferences(IVBE vbe, bool wrapInPvtUDT, (string fieldID, string fieldProperty, bool readOnly) target, IEnumerable<(string fieldID, string fieldProperty, bool readOnly)> fieldIDPairs)
        {
            var refactoredCode = new Dictionary<string, string>();
            (var state, var rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var support = new EncapsulateFieldTestSupport();
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var resolver = support.SetupResolver(state, rewritingManager);
                var encapsulateFieldFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var sutFactory = resolver.Resolve<IEncapsulateFieldReferenceReplacerFactory>();

                var fieldDeclaration = state.DeclarationFinder.MatchName(target.fieldID).Single();

                var fieldCandidate = encapsulateFieldFactory.CreateFieldCandidate(fieldDeclaration);

                if (wrapInPvtUDT)
                {
                    var defaultObjectStateUDT = encapsulateFieldFactory.CreateDefaultObjectStateField(fieldDeclaration.QualifiedModuleName);
                    fieldCandidate = encapsulateFieldFactory.CreateUDTMemberCandidate(fieldCandidate, defaultObjectStateUDT);
                }

                fieldCandidate.PropertyIdentifier = target.fieldProperty;
                fieldCandidate.IsReadOnly = target.readOnly;
                fieldCandidate.EncapsulateFlag = true;

                var sut = sutFactory.Create();
                sut.ReplaceReferences(new IEncapsulateFieldCandidate[] { fieldCandidate }, rewriteSession);

                if (rewriteSession.TryRewrite())
                {
                    refactoredCode = vbe.ActiveVBProject.VBComponents
                        .ToDictionary(component => component.Name, component => component.CodeModule.Content());
                }
            }

            return refactoredCode;
        }
    }
}
