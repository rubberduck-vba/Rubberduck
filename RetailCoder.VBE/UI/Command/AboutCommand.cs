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
}
