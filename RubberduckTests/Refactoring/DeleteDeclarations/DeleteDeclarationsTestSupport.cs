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
        //TODO: Replace ImplementInterface Resource with resource element for this refactoring
        public static string TodoContent => Rubberduck.Resources.Refactorings.Refactorings.ImplementInterface_TODO;

        public IDictionary<string, string> TestRefactoring(Func<RubberduckParserState, IEnumerable<Declaration>> testTargetListBuilder, Func<RubberduckParserState, IEnumerable<Declaration>, IRewritingManager, bool, IExecutableRewriteSession> refactorAction, bool injectTODOComment, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return RefactorCode(vbe, testTargetListBuilder, refactorAction, injectTODOComment);
        }

        private static IDictionary<string, string> RefactorCode(IVBE vbe, Func<RubberduckParserState, IEnumerable<Declaration>> testTargetListBuilder, Func<RubberduckParserState, IEnumerable<Declaration>, IRewritingManager, bool, IExecutableRewriteSession> refactorAction, bool injectTODOComment)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var targets = testTargetListBuilder(state);

                var session = refactorAction(state, targets, rewritingManager, injectTODOComment);

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

        public static DeleteDeclarationsRefactoringAction CreateDeleteDeclarationRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
        {
            var deletionTargetFactory = new DeclarationDeletionTargetFactory(declarationFinderProvider);
            return new DeleteDeclarationsRefactoringAction(declarationFinderProvider,
                 new DeleteModuleElementsRefactoringAction(declarationFinderProvider, deletionTargetFactory, rewritingManager),
                 new DeleteProcedureScopeElementsRefactoringAction(declarationFinderProvider, deletionTargetFactory, rewritingManager),
                 new DeleteUDTMembersRefactoringAction(declarationFinderProvider, deletionTargetFactory, rewritingManager),
                 new DeleteEnumMembersRefactoringAction(declarationFinderProvider, deletionTargetFactory, rewritingManager),
                 rewritingManager);
        }
    }
}
