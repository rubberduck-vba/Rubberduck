using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ChangeFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Refactoring.RenameFolder
{
    [TestFixture]
    public class RenameFolderRefactoringActionTests : RefactoringActionTestBase<RenameFolderModel>
    {
        [Test]
        [Category("Refactorings")]
        public void RenameFolderRefactoringAction_TopLevelFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyNewFolder.MySubFolder""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, RenameFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new RenameFolderModel("MyOldFolder", new List<ModuleDeclaration> { module }, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void RenameFolderRefactoringAction_SubFolder()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyOldFolder.MyOldSubFolder.MyNewFolder""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, RenameFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new RenameFolderModel("MyOldFolder.MyOldSubFolder.SubSub", new List<ModuleDeclaration> { module }, "MyNewFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void RenameFolderRefactoringAction_PreservesSubFolderStructure()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub.Sub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Folder ""MyOldFolder.MyNewFolder.MySubFolder.SubSub.Sub""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, RenameFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new RenameFolderModel("MyOldFolder.MyOldSubFolder", new List<ModuleDeclaration> { module }, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void RenameFolderRefactoringAction_NotInFolder_DoesNothing()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubFolder.SubSub.Sub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = code;

            Func<RubberduckParserState, RenameFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new RenameFolderModel("NotMyOldFolder.MyOldSubFolder", new List<ModuleDeclaration> { module }, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void RenameFolderRefactoringAction_ChangesExactlyTheSpecifiedModules()
        {
            const string code1 = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string code2 = @"
'@Folder(""MyOldFolder.MyOldSubfolder"")
Public Sub Foo()
End Sub
";
            const string code3 = @"
'@Folder(""MyOtherFolder.MyOldSubfolder"")
Public Sub Foo()
End Sub
";
            const string code4 = @"
'@Folder(""MyOtherFolder.MyOtherSubfolder"")
Public Sub Foo()
End Sub
";
            const string code5 = @"
Public Sub Foo()
End Sub
";
            const string expectedCode1 = @"
'@Folder ""MyOldFolder.MyNewFolder.SubSub""
Public Sub Foo()
End Sub
";
            const string expectedCode2 = @"
'@Folder ""MyOldFolder.MyNewFolder""
Public Sub Foo()
End Sub
";
            const string expectedCode3 = code3;
            const string expectedCode4 = code4;
            const string expectedCode5 = code5;
            Func<RubberduckParserState, RenameFolderModel> modelBuilder = (state) =>
            {
                var modules = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .ToList();

                var module1 = modules.Single(module => module.IdentifierName.Equals("SubSubFolderModule"));
                var module2 = modules.Single(module => module.IdentifierName.Equals("SubFolderModuleIncluded"));
                const string originalFolder = "MyOldFolder.MyOldSubfolder";

                return new RenameFolderModel(originalFolder, new List<ModuleDeclaration> { module1, module2 }, "MyNewFolder");
            };

            var refactoredCode = RefactoredCode(
                    modelBuilder,
                ("SubSubFolderModule", code1, ComponentType.StandardModule),
                    ("SubFolderModuleIncluded", code2, ComponentType.ClassModule),
                    ("SubFolderModuleNotIncluded", code3, ComponentType.ClassModule),
                    ("UnaffectedSubFolderModule", code4, ComponentType.StandardModule),
                    ("NoFolderModule", code5, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode1, refactoredCode["SubSubFolderModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["SubFolderModuleIncluded"]);
            Assert.AreEqual(expectedCode3, refactoredCode["SubFolderModuleNotIncluded"]);
            Assert.AreEqual(expectedCode4, refactoredCode["UnaffectedSubFolderModule"]);
            Assert.AreEqual(expectedCode5, refactoredCode["NoFolderModule"]);
        }

        protected override IRefactoringAction<RenameFolderModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var changeFolderAction = new ChangeFolderRefactoringAction(rewritingManager, moveToFolderAction);
            return new RenameFolderRefactoringAction(rewritingManager, changeFolderAction);
        }
    }
}