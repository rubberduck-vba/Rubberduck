using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ChangeFolder;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveFolders
{
    [TestFixture]
    public class MoveFolderRefactoringActionTests : RefactoringActionTestBase<MoveFolderModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveFolderRefactoringAction_NoAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNewFolder.MySubFolder.TestProject""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, MoveFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration; 
                return new MoveFolderModel("TestProject", new List<ModuleDeclaration>{module}, "MyNewFolder.MySubFolder");
            };

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var refactoredCode = RefactoredCode(vbe, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode["TestModule"]);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveFolderRefactoringAction_TopLevelFolder()
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
            Func<RubberduckParserState, MoveFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveFolderModel("MyOldFolder", new List<ModuleDeclaration> { module }, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveFolderRefactoringAction_SubFolder()
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
            Func<RubberduckParserState, MoveFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveFolderModel("MyOldFolder.MyOldSubFolder.SubSub", new List<ModuleDeclaration> { module }, "MyNewFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveFolderRefactoringAction_PreservesSubFolderStructure()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub.Sub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyNewFolder.MySubFolder.MyOldSubFolder.SubSub.Sub""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, MoveFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveFolderModel("MyOldFolder.MyOldSubFolder", new List<ModuleDeclaration> { module }, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveFolderRefactoringAction_WorksForMultipleInFolder()
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

            Func<RubberduckParserState, MoveFolderModel> modelBuilder = (state) =>
            {
                var modules = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .Where(module => module.IdentifierName != "OtherFolderModule")
                    .ToList();
                return new MoveFolderModel("MyOldFolder.MyOldSubFolder", modules, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(
                modelBuilder,
                ("TestModule", code1, ComponentType.StandardModule),
                ("SameFolderModule", code2, ComponentType.StandardModule),
                ("SubFolderModule", code3, ComponentType.StandardModule),
                ("OtherFolderModule", code4, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode1, refactoredCode["TestModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["SameFolderModule"]);
            Assert.AreEqual(expectedCode3, refactoredCode["SubFolderModule"]);
            Assert.AreEqual(expectedCode4, refactoredCode["OtherFolderModule"]);
        }

        protected override IRefactoringAction<MoveFolderModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var changeFolderAction = new ChangeFolderRefactoringAction(rewritingManager, moveToFolderAction);
            return new MoveFolderRefactoringAction(rewritingManager, changeFolderAction);
        }
    }
}