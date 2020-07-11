using System;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveToFolder
{
    [TestFixture]
    public class MoveToFolderRefactoringTests : InteractiveRefactoringTestBase<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_NoAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNewFolder.MySubFolder""

Public Sub Foo()
End Sub
";
            Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var refactoredCode = RefactoredCode(
                "TestModule", 
                DeclarationType.Module, 
                presenterAction,
                null,
                ("TestModule", code, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode, refactoredCode["TestModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_UpdateAnnotation()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyNewFolder.MySubFolder""
Public Sub Foo()
End Sub
";
            Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var refactoredCode = RefactoredCode(
                "TestModule",
                DeclarationType.Module,
                presenterAction,
                null,
                ("TestModule", code, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode, refactoredCode["TestModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_NameContainingDoubleQuotes()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNew""""Folder.My""""""""""""""""SubFolder""

Public Sub Foo()
End Sub
";
            Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNew\"Folder.My\"\"\"\"SubFolder";
                return model;
            };

            var refactoredCode = RefactoredCode(
                "TestModule",
                DeclarationType.Module,
                presenterAction,
                null,
                ("TestModule", code, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode, refactoredCode["TestModule"]);
        }


        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_InitialModel_NoAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var model = InitialModel(vbe, "TestModule", DeclarationType.ProceduralModule);

            var targetName = model.Targets.Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("TestModule", targetName);
            Assert.AreEqual("TestProject", initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_InitialModel_UpdateAnnotation()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule", 
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            var targetName = model.Targets.Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("TestModule", targetName);
            Assert.AreEqual("MyOldFolder.MyOldSubfolder.SubSub", initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_InitialModel_NameContainingDoubleQuotes()
        {
            const string code = @"
'@Folder(""MyNew""""Folder.My""""""""""""""""SubFolder"")
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            var targetName = model.Targets.Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("TestModule", targetName);
            Assert.AreEqual("MyNew\"Folder.My\"\"\"\"SubFolder", initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderRefactoring_TargetNotAModule_Throws()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var refactoredCode = RefactoredCode(
                "Foo",
                DeclarationType.Procedure,
                presenterAction,
                typeof(InvalidDeclarationTypeException),
                ("TestModule", code, ComponentType.StandardModule));
        }


        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel> userInteraction, 
            ISelectionService selectionService)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var moveMultipleToFolderAction = new MoveMultipleToFolderRefactoringAction(rewritingManager, moveToFolderAction);

            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new MoveToFolderRefactoring(moveMultipleToFolderAction, selectedDeclarationProvider, selectionService, userInteraction, state);
        }
    }
}