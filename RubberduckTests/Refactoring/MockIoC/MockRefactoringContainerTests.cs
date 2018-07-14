using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MockIoC
{
    [TestFixture]
    public class MockRefactoringContainerTests
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
        [Apartment(ApartmentState.STA)]
        [Category("MockIoC_Test")]
        public void CanResolve_Actual()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var container = RefactoringContainerInstaller.GetContainer();
                var factory = container.Resolve<IRefactoringPresenterFactory>();

                var model = new RenameModel(state,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = factory.Create<IRenamePresenter, RenameModel>(model);

                Assert.IsInstanceOf<RenamePresenter>(presenter);
            }
        }
    }
}
