namespace Rubberduck.UI.Command.MenuItems
{
    public interface IAppMenu
    {
        void Localize();
        void Initialize();
        void SetCommandButtonEnabledState(string key, bool isEnabled = true);
    }
}