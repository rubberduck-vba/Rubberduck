namespace Rubberduck.UI.Command
{
    public enum RubberduckMenuItemDisplayOrder
    {
        UnitTesting,
        Refactorings,
        Navigate,
        CodeInspections,
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

    public enum NavigationMenuItemDisplayOrder
    {
        CodeExplorer,
        ToDoExplorer,
        FindSymbol,
        FindAllReferences,
        FindImplementations
    }
}