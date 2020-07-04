using System;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ChangeFolder;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveFolders
{
    [TestFixture]
    public class MoveContainingFolderRefactoringTests :InteractiveRefactoringTestBase<IMoveMultipleFoldersPresenter, MoveMultipleFoldersModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_NoAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNewFolder.MySubFolder.TestProject""

Public Sub Foo()
End Sub
";
            Func<MoveMultipleFoldersModel, MoveMultipleFoldersModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var refactoredCode = RefactoredCode(vbe, "TestModule", DeclarationType.ProceduralModule, presenterAction);

            Assert.AreEqual(expectedCode, refactoredCode["TestModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_TopLevelFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldFolder""
Public Sub Foo()
End Sub
";
            Func<MoveMultipleFoldersModel, MoveMultipleFoldersModel> presenterAction = (model) =>
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
        public void MoveContainingFolderRefactoring_SubFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyNewFolder.SubSub""
Public Sub Foo()
End Sub
";
            Func<MoveMultipleFoldersModel, MoveMultipleFoldersModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder";
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
        public void MoveContainingFolderRefactoring_WorksForMultipleInFolder()
        {
            const string code1 = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";
            const string code2 = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";
            const string code3 = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string code4 = @"
'@Folder(""MyOldFolder.MyOtherSubFolder"")
Public Sub Foo()
End Sub
";
            const string expectedCode1 = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldSubFolder""
Public Sub Foo()
End Sub
";
            const string expectedCode2 = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldSubFolder""
Public Sub Foo()
End Sub
";
            const string expectedCode3 = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldSubFolder.SubSub""
Public Sub Foo()
End Sub
";
            const string expectedCode4 = code4;

            Func<MoveMultipleFoldersModel, MoveMultipleFoldersModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var refactoredCode = RefactoredCode(
                "TestModule",
                DeclarationType.Module,
                presenterAction,
                null,
                ("TestModule", code1, ComponentType.StandardModule),
                ("SameFolderModule", code2, ComponentType.StandardModule),
                ("SubFolderModule", code3, ComponentType.StandardModule),
                ("OtherFolderModule", code4, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode1, refactoredCode["TestModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["SameFolderModule"]);
            Assert.AreEqual(expectedCode3, refactoredCode["SubFolderModule"]);
            Assert.AreEqual(expectedCode4, refactoredCode["OtherFolderModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_DistinguishesBetweenSameNameFoldersInDifferentProjects()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";
            const string expectedCode1 = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldSubFolder""
Public Sub Foo()
End Sub
";
            const string expectedCode2 = code;

            Func<MoveMultipleFoldersModel, MoveMultipleFoldersModel> presenterAction = (model) =>
            {
                model.TargetFolder = "MyNewFolder.MySubFolder";
                return model;
            };

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .ProjectBuilder("OtherProject", ProjectProtection.Unprotected)
                .AddComponent("OtherModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var refactoredCode = RefactoredCode(vbe, "TestModule", DeclarationType.ProceduralModule, presenterAction, extractAllProjects: true);

            Assert.AreEqual(expectedCode1, refactoredCode["TestModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["OtherModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_InitialModel_NoAnnotation()
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

            var sourceFolder = model.ModulesBySourceFolder.Keys.Single();
            var targetModuleName = model.ModulesBySourceFolder[sourceFolder].Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("TestProject", sourceFolder);
            Assert.AreEqual("TestModule", targetModuleName);
            Assert.AreEqual(string.Empty, initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_InitialModel_TopLevelFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder"")
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule", 
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            var sourceFolder = model.ModulesBySourceFolder.Keys.Single();
            var targetModuleName = model.ModulesBySourceFolder[sourceFolder].Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("MyOldFolder", sourceFolder);
            Assert.AreEqual("TestModule", targetModuleName);
            Assert.AreEqual(string.Empty, initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_InitialModel_SubFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub"")
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            var sourceFolder = model.ModulesBySourceFolder.Keys.Single();
            var targetModuleName = model.ModulesBySourceFolder[sourceFolder].Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("MyOldFolder.MyOldSubFolder.SubSub", sourceFolder);
            Assert.AreEqual("TestModule", targetModuleName);
            Assert.AreEqual("MyOldFolder.MyOldSubFolder", initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_InitialModel_MultipleInFolder()
        {
            const string code1 = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";
            const string code2 = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";
            const string code3 = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string code4 = @"
'@Folder(""MyOldFolder.MyOtherSubFolder"")
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code1, ComponentType.StandardModule),
                ("SameFolderModule", code2, ComponentType.StandardModule),
                ("SubFolderModule", code3, ComponentType.StandardModule),
                ("OtherFolderModule", code4, ComponentType.StandardModule));

            var sourceFolder = model.ModulesBySourceFolder.Keys.Single();
            var targetModuleNames = model.ModulesBySourceFolder[sourceFolder]
                .Select(module => module.IdentifierName)
                .OrderBy(name => name)
                .ToList();
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("MyOldFolder.MyOldSubFolder", sourceFolder);
            Assert.AreEqual(3, targetModuleNames.Count);
            Assert.AreEqual("SameFolderModule", targetModuleNames[0]);
            Assert.AreEqual("SubFolderModule", targetModuleNames[1]);
            Assert.AreEqual("TestModule", targetModuleNames[2]);
            Assert.AreEqual("MyOldFolder", initialTargetFolder);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveContainingFolderRefactoring_InitialModel_DistinguishesBetweenSameNameFoldersInDifferentProjects()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder"")
Public Sub Foo()
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .ProjectBuilder("OtherProject", ProjectProtection.Unprotected)
                .AddComponent("OtherModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var model = InitialModel(vbe, "TestModule", DeclarationType.ProceduralModule);

            var sourceFolder = model.ModulesBySourceFolder.Keys.Single();
            var targetModuleName = model.ModulesBySourceFolder[sourceFolder].Single().IdentifierName;
            var initialTargetFolder = model.TargetFolder;

            Assert.AreEqual("MyOldFolder.MyOldSubFolder", sourceFolder);
            Assert.AreEqual("TestModule", targetModuleName);
            Assert.AreEqual("MyOldFolder", initialTargetFolder);
        }


        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IMoveMultipleFoldersPresenter, MoveMultipleFoldersModel> userInteraction, 
            ISelectionService selectionService)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var changeFolderAction = new ChangeFolderRefactoringAction(rewritingManager, moveToFolderAction);
            var moveFolderAction = new MoveFolderRefactoringAction(rewritingManager, changeFolderAction);
            var moveMultipleFoldersAction = new MoveMultipleFoldersRefactoringAction(rewritingManager, moveFolderAction);

            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new MoveContainingFolderRefactoring(moveMultipleFoldersAction, selectedDeclarationProvider, selectionService, userInteraction, state, state);
        }
    }
}