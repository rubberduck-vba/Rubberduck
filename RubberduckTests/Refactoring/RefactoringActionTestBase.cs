using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    public abstract class RefactoringActionTestBase<TModel>
        where TModel : class, IRefactoringModel
    {

        protected string RefactoredCode(string code, Func<RubberduckParserState, TModel> modelBuilder)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            var refactored = RefactoredCode(vbe, modelBuilder);
            return refactored[componentName];
        }

        protected IDictionary<string, string> RefactoredCode(Func<RubberduckParserState, TModel> modelBuilder, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return RefactoredCode(vbe, modelBuilder);
        }

        protected IDictionary<string, string> RefactoredCode(IVBE vbe, Func<RubberduckParserState, TModel> modelBuilder)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var baseRefactoring = TestBaseRefactoring(state, rewritingManager);
                var model = modelBuilder(state);

                baseRefactoring.Refactor(model);

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        protected abstract IRefactoringAction<TModel> TestBaseRefactoring(
            RubberduckParserState state,
            IRewritingManager rewritingManager
        );

        protected virtual IVBE TestVbe(string content, out IVBComponent component, Selection? selection = null)
        {
            return MockVbeBuilder.BuildFromSingleStandardModule(content, out component, selection ?? default).Object;
        }

        protected virtual IVBE TestVbe(params (string componentName, string content, ComponentType componentType)[] modules)
        {
            return MockVbeBuilder.BuildFromModules(modules).Object;
        }
    }
}