using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Refactoring.MoveToFolder
{
    [TestFixture]
    public class MoveMultipleToFolderRefactoringActionTests : RefactoringActionTestBase<MoveMultipleToFolderModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveMultipleToFolderBaseRefactoring_Works()
        {
            const string code1 = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string code2 = @"
Public Sub Foo()
End Sub
";
            const string code3 = @"
'@Folder(""MyOldFolder.MyOldSubfolder.SubSub"")
Public Sub Foo()
End Sub
";
            const string code4 = @"
Public Sub Foo()
End Sub
";
            const string expectedCode1 = @"
'@Folder ""MyNewFolder.MySubFolder""
Public Sub Foo()
End Sub
";
            const string expectedCode2 = @"'@Folder ""MyNewFolder.MySubFolder""

Public Sub Foo()
End Sub
";
            const string expectedCode3 = code3;
            const string expectedCode4 = code4;
            Func<RubberduckParserState, MoveMultipleToFolderModel> modelBuilder = (state) =>
            {
                var standardModule = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single(declaration => declaration.IdentifierName.Equals("TestModule")) as ModuleDeclaration;
                var classModule = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .Single(declaration => declaration.IdentifierName.Equals("TestClass")) as ModuleDeclaration;
                return new MoveMultipleToFolderModel(new List<ModuleDeclaration>{standardModule, classModule}, "MyNewFolder.MySubFolder");
            };

            var refactoredCode = RefactoredCode(
                    modelBuilder,
                ("TestModule", code1, ComponentType.StandardModule),
                    ("TestClass", code2, ComponentType.ClassModule),
                    ("OtherClass", code3, ComponentType.ClassModule),
                    ("OtherModule", code4, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode1, refactoredCode["TestModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["TestClass"]);
            Assert.AreEqual(expectedCode3, refactoredCode["OtherClass"]);
            Assert.AreEqual(expectedCode4, refactoredCode["OtherModule"]);
        }

        protected override IRefactoringAction<MoveMultipleToFolderModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            return new MoveMultipleToFolderRefactoringAction(rewritingManager, moveToFolderAction);
        }
    }
}