using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;

namespace RubberduckTests.Refactoring
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

        protected override IRefactoringAction<MoveToFolderModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater();
            return new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
        }
    }
}