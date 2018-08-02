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
    [Category("MockIoC_Test")]
    public class MockRefactoringContainerTests
    {
        [Test]
        public void CanResolve_Mock()
        {
            var container = RefactoringContainerInstaller.GetContainer();
            var presenter = container.Resolve<Mock<IRenamePresenter>>();

            Assert.IsInstanceOf<Mock<IRenamePresenter>>(presenter);
        }

        [Test]
        public void CanResolve_Factory()
        {
            var container = RefactoringContainerInstaller.GetContainer();
            var factory = container.Resolve<IRefactoringPresenterFactory>();

            Assert.IsInstanceOf<IRefactoringPresenterFactory>(factory);
        }

        [Test]
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

                Assert.IsInstanceOf<IRenamePresenter>(presenter);
            }
        }

        [Test]
        public void CanMutateMock_Indirect()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var actual = new RenameModel(state,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(2, 2)));
                var container = RefactoringContainerInstaller.GetContainer();
                var mock = container.Resolve<Mock<IRefactoringView<RenameModel>>>();
                mock.CallBase = true;
                mock.SetupGet(m => m.DataContext).Returns(actual);
                var factory = container.Resolve<IRefactoringPresenterFactory>();

                var model = new RenameModel(state,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = (RenamePresenter)factory.Create<IRenamePresenter, RenameModel>(model);
                var expected = presenter.Dialog.View.DataContext;
                Assert.AreEqual(actual, expected);
            }
        }
    }
}
