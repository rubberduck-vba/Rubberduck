using System.Linq;
using NUnit.Framework;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.UI.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveToFolder
{
    [TestFixture]
    public class MoveMultipleToFolderViewModelTests
    {
        [Test]
        [Category("Refactorings")]
        public void InitialFolderIsInitialTargetFromModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);

                var initialTargetFolder = model.TargetFolder;
                var viewModel = TestViewModel(model);

                Assert.AreEqual(initialTargetFolder, viewModel.NewFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void UpdatingTargetFolderUpdatesModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var viewModel = TestViewModel(model);

                const string newTarget = "Test.Test.Test";
                viewModel.NewFolder = newTarget;

                Assert.AreEqual(newTarget, model.TargetFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(null)]
        [TestCase("")]
        public void EmptyTargetFolder_Error(string emptyFolderName)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var viewModel = TestViewModel(model);

                viewModel.NewFolder = emptyFolderName;

                Assert.IsTrue(viewModel.HasErrors);
                Assert.IsFalse(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase("raeraf afrwefe \n fefaef")]
        [TestCase("raeraf afrwefe \r fefaef")]
        [TestCase("raeraf afrwefe \u0000 fefaef")]
        public void TargetFolderWithControlCharacter_Error(string folderName)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var viewModel = TestViewModel(model);

                viewModel.NewFolder = folderName;

                Assert.IsTrue(viewModel.HasErrors);
                Assert.IsFalse(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(".SomeFolder.SomeOtherFolder")]
        [TestCase("SomeFolder..SomeOtherFolder")]
        [TestCase("SomeFolder.SomeOtherFolder.")]
        public void TargetFolderWithEmptyIndividualFolder_Error(string folderName)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var viewModel = TestViewModel(model);

                viewModel.NewFolder = folderName;

                Assert.IsTrue(viewModel.HasErrors);
                Assert.IsFalse(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void NonEmptyTargetFolderWithoutControlCharacter_NoError()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo" , state.DeclarationFinder);
                var viewModel = TestViewModel(model);

                viewModel.NewFolder = ";oehaha .adaiafa.a@#$^%&#@$&%%$%^$.ad3.1010101.  ## . @.{ ]. rqrq";

                Assert.IsFalse(viewModel.HasErrors);
                Assert.IsTrue(viewModel.IsValidFolder);
            }
        }

        private MoveMultipleToFolderViewModel TestViewModel(MoveMultipleToFolderModel model)
        {
            return new MoveMultipleToFolderViewModel(model);
        }

        private MoveMultipleToFolderModel TestModel(string sourceFolder, DeclarationFinder finder)
        {
            var modulesToMove = finder.UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Where(module => module.CustomFolder.Equals(sourceFolder)
                                 || module.CustomFolder.IsSubFolderOf(sourceFolder))
                .ToList();

            var initialTarget = sourceFolder;

            return new MoveMultipleToFolderModel(modulesToMove, initialTarget);
        }

        private IVBE TestVbe()
        {
            const string targetFolderComponentCode = @"
'@Folder ""Test.Foo.Bar.Test.Baz""";

            const string component1Code = @"
'@Folder ""FooBar.Foo.Barr.Foo.Test""";

            const string component2Code = @"
'@Folder ""FooBar.Foo.Barz.Test.Foo""";

            return new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TargetFolderComponent", ComponentType.ClassModule, targetFolderComponentCode)
                .AddComponent("Component1", ComponentType.ClassModule, component1Code)
                .AddComponent("Component2", ComponentType.ClassModule, component2Code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;
        }
    }
}