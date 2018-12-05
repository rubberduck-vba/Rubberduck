using Moq;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.Interaction;
using static RubberduckTests.Refactoring.Rename.RenameTestExecution;
using System;
using Rubberduck.Parsing.Rewriter;

namespace RubberduckTests.Refactoring.Rename
{
    internal class RenameTestsDataObject
    {
        public RenameTestsDataObject(string selection, string newName)
        {
            ProjectName = "TestProject";
            MsgBoxReturn = ConfirmationOutcome.Yes;
            RefactorParamType = RefactorParams.QualifiedSelection;
            RawSelection = null;
            NewName = newName;
            OriginalName = selection;
            ModuleTestSetupDefs = new List<RenameTestModuleDefinition>();
            RenameRefactoringUnderTest = null;
            UseLibraries = false;
        }

        public IVBE VBE { get; set; }
        public RubberduckParserState ParserState { get; set; }
        public IRewritingManager RewritingManager { get; set; }
        public string ProjectName { get; set; }
        public string NewName { get; set; }
        public string SelectionModuleName { get; set; }
        public QualifiedSelection QualifiedSelection { get; set; }
        public RenameModel RenameModel { get; set; }
        public Mock<IMessageBox> MsgBox { get; set; }
        public ConfirmationOutcome MsgBoxReturn { get; set; }
        public RefactorParams RefactorParamType { get; set; }
        public Selection? RawSelection { get; set; }
        public List<RenameTestModuleDefinition> ModuleTestSetupDefs { get; set; }
        public string OriginalName { get; set; }
        public RenameRefactoring RenameRefactoringUnderTest { get; set; }
        public Action<RenameTestsDataObject> AdditionalSetup { get; set; }
        public bool UseLibraries { get; set; }
    }
    
}
