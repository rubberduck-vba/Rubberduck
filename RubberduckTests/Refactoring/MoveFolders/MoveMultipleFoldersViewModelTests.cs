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
using Rubberduck.UI.Refactorings.MoveFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveFolders
{
    [TestFixture]
    public class MoveMultipleFoldersViewModelTests
    {
        [Test]
        [Category("Refactorings")]
        public void InitialFolderIsInitialTargetFromModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> {"FooBar.Foo.Barr"}, state.DeclarationFinder);

                var initialTargetFolder = model.TargetFolder;
                var viewModel = TestViewModel(model, state, null);

                Assert.AreEqual(initialTargetFolder, viewModel.NewFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void UpdatingTargetFolderUpdatesModel()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr" }, state.DeclarationFinder);
                var viewModel = TestViewModel(model, state, null);

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
                var model = TestModel(new List<string> { "FooBar.Foo.Barr" }, state.DeclarationFinder);
                var viewModel = TestViewModel(model, state, null);

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
                var model = TestModel(new List<string> { "FooBar.Foo.Barr" }, state.DeclarationFinder);
                var viewModel = TestViewModel(model, state, null);

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
                var model = TestModel(new List<string> { "FooBar.Foo.Barr" }, state.DeclarationFinder);
                var viewModel = TestViewModel(model, state, null);

                viewModel.NewFolder = folderName;

                Assert.IsTrue(viewModel.HasErrors);
                Assert.IsFalse(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void TargetFolderWithoutEmptyPartsOrControlCharacter_NoError()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr" }, state.DeclarationFinder);
                var viewModel = TestViewModel(model, state, null);

                viewModel.NewFolder = ";oehaha .adaiafa.a@#$^%&#@$&%%$%^$.ad3.1010101.  ## . @.{ ]. rqrq";

                Assert.IsFalse(viewModel.HasErrors);
                Assert.IsTrue(viewModel.IsValidFolder);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void SameNameSourceFolders_AsksForConfirmation()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo", "FooBar.Foo.Bar.Test.Foo" }, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> {true});

                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolder = "Test.Foo.Bar.Test";

                viewModel.OkButtonCommand.Execute(null);

                messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), true), Times.Once);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(true, RefactoringDialogResult.Execute)]
        [TestCase(false, RefactoringDialogResult.Cancel)]
        public void SameNameSourceFolders_ResultBasedOnConfirmation(bool confirms, RefactoringDialogResult expectedResult)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo", "FooBar.Foo.Bar.Test.Foo" }, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> { confirms });
                var viewModel = TestViewModel(model, state, messageBoxMock.Object);
                
                viewModel.NewFolder = "Test.Foo.Bar.Test";

                var executionResults = new List<RefactoringDialogResult>();
                void DialogEventCloseHandler(object sender, RefactoringDialogResult result) => executionResults.Add(result);

                viewModel.OnWindowClosed += DialogEventCloseHandler;
                viewModel.OkButtonCommand.Execute(null);
                viewModel.OnWindowClosed -= DialogEventCloseHandler;

                var actualResult = executionResults.Single();

                Assert.AreEqual(expectedResult, actualResult);
            }
        }

        [Test]
        [Category("Refactorings")]
        public void SameNameTargetSubFolder_AsksForConfirmation()
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo"}, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> { true });

                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolder = "Test";

                viewModel.OkButtonCommand.Execute(null);

                messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), true), Times.Once);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(true, RefactoringDialogResult.Execute)]
        [TestCase(false, RefactoringDialogResult.Cancel)]
        public void SameNameTargetSubFolder_ResultBasedOnConfirmation(bool confirms, RefactoringDialogResult expectedResult)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo"}, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> { confirms });
                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolder = "Test";

                var executionResults = new List<RefactoringDialogResult>();
                void DialogEventCloseHandler(object sender, RefactoringDialogResult result) => executionResults.Add(result);

                viewModel.OnWindowClosed += DialogEventCloseHandler;
                viewModel.OkButtonCommand.Execute(null);
                viewModel.OnWindowClosed -= DialogEventCloseHandler;

                var actualResult = executionResults.Single();

                Assert.AreEqual(expectedResult, actualResult);
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(true, true, 2)]
        [TestCase(false, false, 1)]
        [TestCase(true, false, 2)]
        [TestCase(false, true, 1)]
        public void SameNameSourceFoldersAndSameNameTargetSubFolder_AsksForConfirmationForBoth(bool confirmsFirst, bool confirmsSecond, int timesExpectedToAsk)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo", "FooBar.Foo.Bar.Test.Foo" }, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> { confirmsFirst, confirmsSecond });

                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolder = "Test";

                viewModel.OkButtonCommand.Execute(null);

                messageBoxMock.Verify(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), true), Times.Exactly(timesExpectedToAsk));
            }
        }

        [Test]
        [Category("Refactorings")]
        [TestCase(true, true, RefactoringDialogResult.Execute)]
        [TestCase(false, false, RefactoringDialogResult.Cancel)]
        [TestCase(true, false, RefactoringDialogResult.Cancel)]
        [TestCase(false, true, RefactoringDialogResult.Cancel)]
        public void SameNameSourceFoldersAndSameNameTargetSubFolder_ResultBasedOnConfirmation(bool confirmsFirst, bool confirmsSecond, RefactoringDialogResult expectedResult)
        {
            using (var state = MockParser.CreateAndParse(TestVbe()))
            {
                var model = TestModel(new List<string> { "FooBar.Foo.Barr.Foo", "FooBar.Foo.Bar.Test.Foo" }, state.DeclarationFinder);
                var messageBoxMock = MessageBoxMock(new List<bool> { confirmsFirst, confirmsSecond });
                var viewModel = TestViewModel(model, state, messageBoxMock.Object);

                viewModel.NewFolder = "Test";

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

        private MoveMultipleFoldersViewModel TestViewModel(MoveMultipleFoldersModel model, IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox)
        {
            return new MoveMultipleFoldersViewModel(model, messageBox, declarationFinderProvider);
        }

        private MoveMultipleFoldersModel TestModel(ICollection<string> sourceFolders, DeclarationFinder finder)
        {
            var modulesBySourceFolder = new Dictionary<string, ICollection<ModuleDeclaration>>();

            foreach (var sourceFolder in sourceFolders.Distinct())
            {
                modulesBySourceFolder[sourceFolder] = finder.UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .Where(module => module.CustomFolder.Equals(sourceFolder)
                                     || module.CustomFolder.IsSubFolderOf(sourceFolder))
                    .ToList();
            }

            var initialTarget = sourceFolders.FirstOrDefault()?.ParentFolder();

            return new MoveMultipleFoldersModel(modulesBySourceFolder, initialTarget);
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