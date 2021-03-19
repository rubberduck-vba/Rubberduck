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
using Rubberduck.Parsing.Symbols;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class ReferenceReplacerTestSupport
    {
        public static IDictionary<string, string> TestReferenceReplacement(bool wrapInPrivateUDT, (string, string, bool) testTargetTuple, params (string, string, ComponentType)[] moduleTuples)
        {
            var vbe = MockVbeBuilder.BuildFromModules(moduleTuples);
            return ReplaceReferences(vbe.Object, wrapInPrivateUDT, testTargetTuple);
        }

        private static IDictionary<string, string> ReplaceReferences(IVBE vbe, bool wrapInPvtUDT, (string fieldID, string propertyOrUDTMemberID, bool readOnly) target, params (string fieldID, string propertyOrUDTMemberID, bool readOnly)[] fieldTuples)
            => ReplaceReferences(vbe, wrapInPvtUDT, target, fieldTuples.ToList());

        private static IDictionary<string, string> ReplaceReferences(IVBE vbe, bool wrapInPvtUDT, (string fieldID, string propertyOrUDTMemberID, bool readOnly) target, IEnumerable<(string fieldID, string propertyOrUDTMemberID, bool readOnly)> fieldTuples)
        {
            var refactoredCode = new Dictionary<string, string>();
            (var state, var rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var support = new EncapsulateFieldTestSupport();
                var resolver = support.SetupResolver(state, rewritingManager);

                var encapsulateFieldFactory = resolver.Resolve<IEncapsulateFieldCandidateFactory>();
                var fieldCandidate = encapsulateFieldFactory.CreateFieldCandidate(state.DeclarationFinder.MatchName(target.fieldID).Single());

                if (wrapInPvtUDT)
                {
                    var defaultObjectStateUDT = encapsulateFieldFactory.CreateDefaultObjectStateField(fieldCandidate.QualifiedModuleName);
                    fieldCandidate = encapsulateFieldFactory.CreateUDTMemberCandidate(fieldCandidate, defaultObjectStateUDT);
                }

                //For ReferenceReplacer tests, UDTMember identifiers == PropertyIdentifiers
                fieldCandidate.PropertyIdentifier = target.propertyOrUDTMemberID;
                fieldCandidate.IsReadOnly = target.readOnly;
                fieldCandidate.EncapsulateFlag = true;

                var selected = new IEncapsulateFieldCandidate[] { fieldCandidate };
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var sut = resolver.Resolve<IEncapsulateFieldReferenceReplacerFactory>().Create();
                sut.ReplaceReferences(selected, rewriteSession);

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
