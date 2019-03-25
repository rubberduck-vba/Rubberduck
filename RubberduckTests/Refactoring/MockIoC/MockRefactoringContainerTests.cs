using System.Collections.Generic;
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

                var model = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = factory.Create<IRenamePresenter, RenameModel>(model);

                Assert.IsInstanceOf<IRenamePresenter>(presenter);
            }
        }

        [Test]
        public void CanMutateMock_Indirect_View()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var actual = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(2, 2)));
                var initial = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(3, 3)));

                var container = RefactoringContainerInstaller.GetContainer();
                var mockArgs =
                    new Dictionary<string, object>
                    {
                        {"behavior", MockBehavior.Default},
                        {"args", new object[] {initial}}
                    };
                var mockView = container.Resolve<Mock<RefactoringViewStub<RenameModel>>>(mockArgs);
                mockView.CallBase = true;
                mockView.SetupGet(m => m.DataContext).Returns(actual);

                var factory = container.Resolve<IRefactoringPresenterFactory>();
                var model = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = (RenamePresenter)factory.Create<IRenamePresenter, RenameModel>(model);
                var expected = presenter.Dialog.View.DataContext;
                Assert.AreEqual(actual, expected);
            }
        }

        [Test]
        public void CanMutateMock_Direct_Dialog()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var actual = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(2, 2)));
                var container = RefactoringContainerInstaller.GetContainer();
                var factory = container.Resolve<IRefactoringPresenterFactory>();

                var model = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = (RenamePresenter)factory.Create<IRenamePresenter, RenameModel>(model);

                var mock = Mock.Get(presenter.Dialog);
                mock.SetupGet(m => m.Model).Returns(actual);

                var expected = presenter.Dialog.Model;
                Assert.AreEqual(actual, expected);
            }
        }

        [Test]
        public void CanMutateMock_Indirect_Dialog()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var actual = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(2, 2)));
                var container = RefactoringContainerInstaller.GetContainer();
                var factory = container.Resolve<IRefactoringPresenterFactory>();

                var model = new RenameModel(state.DeclarationFinder,
                    new QualifiedSelection(new QualifiedModuleName(component), new Selection(1, 1)));
                var presenter = (RenamePresenter)factory.Create<IRenamePresenter, RenameModel>(model);
                
                //Mock setup must happen after creating the presenter. Otherwise, the code will error about 
                //lacking a parameterless constructor since this Resolve will not have the args the stub needs.
                //Also note that the generic parameter must be exactly the same; otherwise we get a different
                //mock object which will result in a test failure.
                var mock = container.Resolve<Mock<RefactoringDialogStub<RenameModel, IRefactoringView<RenameModel>, IRefactoringViewModel<RenameModel>>>>();
                mock.SetupGet(m => m.Model).Returns(actual);

                var expected = presenter.Dialog.Model;
                Assert.AreEqual(actual, expected);
            }
        }
    }
}
