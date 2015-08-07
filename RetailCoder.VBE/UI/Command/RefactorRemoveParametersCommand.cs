namespace Rubberduck.UI.Command
{
    public class RefactorRemoveParametersCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return RubberduckUI.RefactorMenu_RemoveParameter; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RemoveParameters; } }
    }
}