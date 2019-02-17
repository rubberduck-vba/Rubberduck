using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public abstract class RefactoringTestBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void NoActiveSelection_Throws()
        {
            var rewritingManager = new Mock<IRewritingManager>().Object;
            var refactoring = TestRefactoring(rewritingManager, null, null);

            Assert.Throws<NoActiveSelectionException>(() => refactoring.Refactor());
        }

        protected abstract IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            QualifiedSelection? initialSelection = null
        );
        
        protected static ISelectionService MockedSelectionService(QualifiedSelection? initialSelection)
        {
            QualifiedSelection? activeSelection = null;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }
    }
}