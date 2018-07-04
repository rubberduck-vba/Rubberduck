using Moq;
using NUnit.Framework;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;

namespace RubberduckTests.Refactoring.MockIoC
{
    [TestFixture]
    public class MockIocTests
    {
        [Test]
        [Category("MockIoC_Test")]
        public void CanResolve_Mock()
        {
            var container = RefactoringContainerInstaller.GetContainer();
            var presenter = container.Resolve<Mock<IRenamePresenter>>();

            Assert.IsInstanceOf<Mock<IRenamePresenter>>(presenter);
        }

        [Test]
        [Category("MockIoC_Test")]
        public void CanResolve_Factory()
        {
            var container = RefactoringContainerInstaller.GetContainer();
            var factory = container.Resolve<IRefactoringPresenterFactory>();

            Assert.IsInstanceOf<IRefactoringPresenterFactory>(factory);
        }

        [Test]
        [Category("MockIoC_Test")]
        public void CanResolve_Actual()
        {
            var container = RefactoringContainerInstaller.GetContainer();
            var presenter = container.Resolve<IRenamePresenter>();

            Assert.IsInstanceOf<RenamePresenter>(presenter);
        }
    }
}
