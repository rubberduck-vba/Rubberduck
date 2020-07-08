using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveToFolder
{
    [TestFixture]
    public class MoveToFolderRefactoringActionTests : RefactoringActionTestBase<MoveToFolderModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveToFolderBaseRefactoring_NoAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNewFolder.MySubFolder""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, MoveToFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration; 
                return new MoveToFolderModel(module, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderBaseRefactoring_UpdateAnnotation()
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
            Func<RubberduckParserState, MoveToFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveToFolderModel(module, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderBaseRefactoring_NameContainingDoubleQuotes()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNew""""Folder.My""""""""""""""""SubFolder""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, MoveToFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveToFolderModel(module, "MyNew\"Folder.My\"\"\"\"SubFolder");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderBaseRefactoring_HasAnnotation_SameFolder_DoesNotDoAnything()
        {
            const string code = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string expectedCode = code;
            Func<RubberduckParserState, MoveToFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveToFolderModel(module, "MyOldFolder.MyOldSubfolder.SubSub");
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void MoveToFolderBaseRefactoring_NoAnnotation_SameFolder_AddsAnnotation()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""TestProject""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, MoveToFolderModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single() as ModuleDeclaration;
                return new MoveToFolderModel(module, "TestProject");
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

        protected override IRefactoringAction<MoveToFolderModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            return new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
        }
    }
}