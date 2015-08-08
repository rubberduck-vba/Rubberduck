namespace Rubberduck.UI.Command
{
    public enum RubberduckMenuItemDisplayOrder
    {
        UnitTesting,
        Refactorings,
        Navigate,
        CodeInspections,
        CodeExplorer,
        ToDoExplorer,
        SourceControl,
        Options,
        About
    }

    public enum UnitTestingMenuItemDisplayOrder
    {
        TestExplorer,
        RunAllTests
    }

    public enum RefactoringsMenuItemDisplayOrder
    {
        ExtractMethod,
        RenameIdentifier,
        ReorderParameters,
        RemoveParameters
    }
}