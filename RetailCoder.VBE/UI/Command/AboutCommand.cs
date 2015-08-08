namespace Rubberduck.UI.Command
{
    public class AboutCommand : ICommand
    {
        public void Execute()
        {
            using (var window = new AboutWindow())
            {
                window.ShowDialog();
            }
        }
    }

    public class AboutCommandMenuItem : CommandMenuItemBase
    {
        public AboutCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_About"; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.About; } }
    }
}
