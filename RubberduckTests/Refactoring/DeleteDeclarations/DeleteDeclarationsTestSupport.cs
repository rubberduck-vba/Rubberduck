using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
    internal class DeleteDeclarationsTestSupport
    {
        internal List<string> GetRetainedLines(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
            => GetRetainedCodeBlock(moduleCode, modelBuilder)
                .Trim()
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .ToList();

        internal string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
        {
            var refactoredCode = ModifiedCode(
                modelBuilder,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        public IDictionary<string, string> ModifiedCode(Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return ModifiedCode(vbe, modelBuilder);
        }

        private static IDictionary<string, string> ModifiedCode(IVBE vbe, Func<RubberduckParserState, IEnumerable<Declaration>> modelBuilder)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var refactoringAction = CreateDeleteDeclarationRefactoringAction(state, rewritingManager);

                var session = rewritingManager.CheckOutCodePaneSession();

                var targets =  modelBuilder(state);

                var model = new DeleteDeclarationsModel(targets);

                refactoringAction.Refactor(model, session);

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

        public static DeleteDeclarationsRefactoringAction CreateDeleteDeclarationRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
        {
            return new DeleteDeclarationsRefactoringAction(declarationFinderProvider,
                 new DeleteModuleElementsRefactoringAction(declarationFinderProvider, rewritingManager),
                 new DeleteProcedureScopeElementsRefactoringAction(declarationFinderProvider, rewritingManager),
                 new DeleteUDTMembersRefactoringAction(declarationFinderProvider, rewritingManager),
                 new DeleteEnumMembersRefactoringAction(declarationFinderProvider, rewritingManager),
                 rewritingManager);
        }
    }
}
