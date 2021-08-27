using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    public class DeleteDeclarationsTestSupport
    {
        public static string TodoContent => Rubberduck.Resources.Refactorings.Refactorings.CommentVerification_TODO;

        public IDictionary<string, string> TestRefactoring(Func<RubberduckParserState, IEnumerable<Declaration>> testTargetListBuilder, Func<RubberduckParserState, IEnumerable<Declaration>, IRewritingManager, Action<IDeleteDeclarationsModel>, IExecutableRewriteSession> refactorAction, Action<IDeleteDeclarationsModel> modelFlagsAction, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return RefactorCode(vbe, testTargetListBuilder, refactorAction, modelFlagsAction);
        }

        public void DefaultModelFlagAction(IDeleteDeclarationsModel model)
        {
            model.InsertValidationTODOForRetainedComments = false;
        }

        private static IDictionary<string, string> RefactorCode(IVBE vbe, Func<RubberduckParserState, IEnumerable<Declaration>> testTargetListBuilder, Func<RubberduckParserState, IEnumerable<Declaration>, IRewritingManager, Action<IDeleteDeclarationsModel>, IExecutableRewriteSession> refactorAction, Action<IDeleteDeclarationsModel> modelFlagsAction)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var targets = testTargetListBuilder(state);

                var session = refactorAction(state, targets, rewritingManager, modelFlagsAction);

                session.TryRewrite();

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        internal IEnumerable<Declaration> TestTargets(IDeclarationFinderProvider declarationFinderProvider, params string[] targetIdentifiers)
        {
            var targetsList = new List<Declaration>();
            foreach (var tgt in targetIdentifiers)
            {
                targetsList.Add(declarationFinderProvider.DeclarationFinder.MatchName(tgt).Single());
            }

            return targetsList;
        }

        internal IEnumerable<Declaration> TestTargetsUsingDeclarationType(IDeclarationFinderProvider declarationFinderProvider, params (string, DeclarationType)[] targetIdentifiersAccessorPair)
        {
            var targetsList = new List<Declaration>();
            foreach ((string id, DeclarationType decType) in targetIdentifiersAccessorPair)
            {
                var target = declarationFinderProvider.DeclarationFinder
                    .MatchName(id)
                    .Single(t => t.DeclarationType == decType);

                targetsList.Add(target);
            }

            return targetsList;
        }
        internal IEnumerable<Declaration> TestTargetsUsingParentDeclaration(IDeclarationFinderProvider declarationFinderProvider, params (string, string)[] targetIdentifiersAccessorPair)
        {
            var targetsList = new List<Declaration>();
            foreach ((string id, string parentID) in targetIdentifiersAccessorPair)
            {
                var parentDeclaration = declarationFinderProvider.DeclarationFinder
                    .MatchName(parentID).Single();

                var target = declarationFinderProvider.DeclarationFinder
                    .MatchName(id)
                    .Single(t => t.ParentDeclaration == parentDeclaration);

                targetsList.Add(target);
            }

            return targetsList;
        }

        public void SetupAndInvokeIDeclarationDeletionTargetTest(string inputCode, string targetIdentifier, Action<IDeclarationDeletionTarget> testSUT)
        {
            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule)).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName(targetIdentifier).First();

                var resolver = new DeleteDeclarationsTestsResolver(state, rewritingManager);

                var deleteTarget = resolver.Resolve<IDeclarationDeletionTargetFactory>()
                    .Create(target, rewritingManager.CheckOutCodePaneSession());

                testSUT(deleteTarget);
            }
        }
    }
}
