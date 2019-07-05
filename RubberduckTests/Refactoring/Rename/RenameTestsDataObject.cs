using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System;

namespace RubberduckTests.Refactoring.Rename
{
    internal class RenameTestsDataObject
    {
        public RenameTestsDataObject(string declarationName, DeclarationType declarationType, string newName)
        {
            RefactorParamType = RenameTests.RefactorParams.Declaration;
            TargetDeclarationName = declarationName;
            TargetDeclarationType = declarationType;
            NewName = newName;

            ModuleTestSetupDefs = new List<RenameTestModuleDefinition>();
        }

        public RenameTestsDataObject(string selectedIdentifier, string newName)
        {
            RefactorParamType = RenameTests.RefactorParams.QualifiedSelection;
            NewName = newName;
            IntendedSelection = selectedIdentifier;
            ModuleTestSetupDefs = new List<RenameTestModuleDefinition>();
            UseLibraries = false;
        }

        public RenameTestsDataObject(string selectedModule, Selection selection, string newName)
        {
            RefactorParamType = RenameTests.RefactorParams.QualifiedSelection;
            SelectionModuleName = selectedModule;
            RawSelection = selection;
            NewName = newName;

            ModuleTestSetupDefs = new List<RenameTestModuleDefinition>();
        }

        public string ProjectName = "TestProject";

        public IVBE VBE { get; set; }
        public string NewName { get; set; }
        public string SelectionModuleName { get; set; }
        public QualifiedSelection QualifiedSelection { get; set; }
        public string TargetDeclarationName { get; set; }
        public DeclarationType TargetDeclarationType { get; set; }
        public bool DoNotRename { get; set; }
        public Func<RenameModel, RenameModel> PresenterAdjustmentAction { get; set; }
        public RenameTests.RefactorParams RefactorParamType { get; set; }
        public Selection? RawSelection { get; set; }
        public List<RenameTestModuleDefinition> ModuleTestSetupDefs { get; set; }
        public string IntendedSelection { get; set; }
        public Action<RenameTestsDataObject> AdditionalSetup { get; set; }
        public bool UseLibraries { get; set; }

        //Test results
        public RenameModel Model { get; set; }
        public Type ExpectedException { get; set; }
        public IDictionary<string, string> ActualCode { get; set; }
    }
    
}
