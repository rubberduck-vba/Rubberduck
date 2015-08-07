namespace Rubberduck.UI.Command
{
    public class RefactorReorderParametersCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorReorderParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorReorderParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return RubberduckUI.RefactorMenu_ReorderParameters; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ReorderParameters; } }
    }
}