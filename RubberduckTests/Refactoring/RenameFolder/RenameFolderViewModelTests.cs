using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.UI.Refactorings.MoveFolder;
using Rubberduck.UI.Refactorings.MoveToFolder;
using Rubberduck.UI.Refactorings.RenameFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.RenameFolder
{
    [TestFixture]
    public class RenameFolderViewModelTests
    {
        [Test]
        [Category("Refactorings")]
        public void InitialNewFolderNameIsNewSubFolderNameFromModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);

                var initialNewSubfolderName = model.NewSubFolder;
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                Assert.AreEqual(initialNewSubfolderName, viewModel.NewFolderName);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void UpdatingNewFolderUpdatesModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                const string newSubFolder = "Test.Test.Test";
                viewModel.NewFolderName = newSubFolder;

                Assert.AreEqual(newSubFolder, model.NewSubFolder);
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
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                viewModel.NewFolderName = emptyFolderName;

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
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                viewModel.NewFolderName = folderName;

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
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                viewModel.NewFolderName = folderName;

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
                var model = TestModel("FooBar.Foo", state.DeclarationFinder);
                var messageBox = MessageBoxMock(new List<bool>()).Object;
                var viewModel = TestViewModel(model, state, messageBox);

                viewModel.NewFolderName = ";oehaha .adaiafa.a@#$^%&#@$&%%$%^$.ad3.1010101.  ## . @.{ ]. rqrq";

                Assert.IsFalse(viewModel.HasErrors);
                Assert.IsTrue(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void FolderAlreadyExists_AsksForConfirmation()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo.Barr", state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool>());
                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolderName = "Barz";

                viewModel.OkButtonCommand.Execute(null);

                messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), true), Times.Once);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(true, RefactoringDialogResult.Execute)]
        [TestCase(false, RefactoringDialogResult.Cancel)]
        public void FolderAlreadyExists_ResultBasedOnConfirmation(bool confirms, RefactoringDialogResult expectedResult)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel("FooBar.Foo.Barr", state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool>{ confirms });
                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolderName = "Barz";

                var executionResults = new List<RefactoringDialogResult>();
                void DialogEventCloseHandler(object sender, RefactoringDialogResult result) => executionResults.Add(result);

                viewModel.OnWindowClosed += DialogEventCloseHandler;
                viewModel.OkButtonCommand.Execute(null);
                viewModel.OnWindowClosed -= DialogEventCloseHandler;

                var actualResult = executionResults.Single();

                Assert.AreEqual(expectedResult, actualResult);
            }
        }

        private Mock<IMessageBox> MessageBoxMock(IList<bool> confirmsRequest)
        {
            var requestCounter = 0;
            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()))
                .Returns<string, string, bool>((message, caption, suggestion) => requestCounter < confirmsRequest.Count
                                                                                 && confirmsRequest[requestCounter++]);
            return messageBox;
        }

        private RenameFolderViewModel TestViewModel(RenameFolderModel model, IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox)
        {
            return new RenameFolderViewModel(model, messageBox, declarationFinderProvider);
        }

        private RenameFolderModel TestModel(string sourceFolder, DeclarationFinder finder)
        {
            var modulesToMove = finder.UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Where(module => module.CustomFolder.Equals(sourceFolder)
                                 || module.CustomFolder.IsSubFolderOf(sourceFolder))
                .ToList();

            var initialTarget = sourceFolder.SubFolderName();

            return new RenameFolderModel(sourceFolder, modulesToMove, initialTarget);
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