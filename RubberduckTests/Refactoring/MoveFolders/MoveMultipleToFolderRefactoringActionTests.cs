using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ChangeFolder;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Refactoring.MoveFolders
{
    [TestFixture]
    public class MoveMultipleFoldersRefactoringActionTests : RefactoringActionTestBase<MoveMultipleFoldersModel>
    {
        [Test]
        [Category("Refactorings")]
        public void MoveMultipleFoldersRefactoringAction_WorksForMultipleFolders()
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
'@Folder(""MyOtherFolder.MyOtherOldSubfolder"")
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
'@Folder ""MyNewFolder.MyOldSubfolder.SubSub""
Public Sub Foo()
End Sub
";
            const string expectedCode2 = @"
'@Folder ""MyNewFolder.MyOldSubfolder""
Public Sub Foo()
End Sub
";
            const string expectedCode3 = @"
'@Folder ""MyNewFolder.MyOtherOldSubfolder""
Public Sub Foo()
End Sub
";
            const string expectedCode4 = code4;
            const string expectedCode5 = code5;
            Func<RubberduckParserState, MoveMultipleFoldersModel> modelBuilder = (state) =>
            {
                var modules = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .OfType<ModuleDeclaration>()
                    .ToList();

                var firstFolderModules = modules
                    .Where(module => module.CustomFolder.Equals("MyOldFolder.MyOldSubfolder")
                                     || module.CustomFolder.IsSubFolderOf("MyOldFolder.MyOldSubfolder"))
                    .ToList();

                var secondFolderModules = modules
                    .Where(module => module.CustomFolder.Equals("MyOtherFolder.MyOtherOldSubfolder")
                                     || module.CustomFolder.IsSubFolderOf("MyOtherFolder.MyOtherOldSubfolder"))
                    .ToList();

                var modulesByFolders = new Dictionary<string, ICollection<ModuleDeclaration>>
                {
                    {"MyOldFolder.MyOldSubfolder", firstFolderModules},
                    {"MyOtherFolder.MyOtherOldSubfolder", secondFolderModules}
                };

                return new MoveMultipleFoldersModel(modulesByFolders, "MyNewFolder");
            };

            var refactoredCode = RefactoredCode(
                    modelBuilder,
                ("SubSubFolderModule", code1, ComponentType.StandardModule),
                    ("SubFolderModule", code2, ComponentType.ClassModule),
                    ("OtherSubFolderModule", code3, ComponentType.ClassModule),
                    ("UnaffectedSubFolderModule", code4, ComponentType.StandardModule),
                    ("NoFolderModule", code5, ComponentType.StandardModule));

            Assert.AreEqual(expectedCode1, refactoredCode["SubSubFolderModule"]);
            Assert.AreEqual(expectedCode2, refactoredCode["SubFolderModule"]);
            Assert.AreEqual(expectedCode3, refactoredCode["OtherSubFolderModule"]);
            Assert.AreEqual(expectedCode4, refactoredCode["UnaffectedSubFolderModule"]);
            Assert.AreEqual(expectedCode5, refactoredCode["NoFolderModule"]);
        }

        protected override IRefactoringAction<MoveMultipleFoldersModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var changeFolderAction = new ChangeFolderRefactoringAction(rewritingManager, moveToFolderAction);
            var moveFolderAction = new MoveFolderRefactoringAction(rewritingManager, changeFolderAction);
            return new MoveMultipleFoldersRefactoringAction(rewritingManager, moveFolderAction);
        }
    }
}