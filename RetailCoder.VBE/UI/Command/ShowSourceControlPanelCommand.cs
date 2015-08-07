namespace Rubberduck.UI.Command
{
    public class ShowSourceControlPanelCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class ShowSourceControlPanelCommandMenuItem : CommandMenuItemBase
    {
        public ShowSourceControlPanelCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return RubberduckUI.RubberduckMenu_SourceControl; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.SourceControl; } }
    }
}