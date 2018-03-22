using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Implement this interface in ViewModel classes that contain navigatable items.
    /// </summary>
    public interface INavigateSelection
    {
        INavigateSource SelectedItem { get; }
        INavigateCommand NavigateCommand { get; }
    }

    /// <summary>
    /// Implement this interface in ViewModel classes that can be double-click-navigated.
    /// </summary>
    public interface INavigateSource
    {
        NavigateCodeEventArgs GetNavigationArgs();
    }
}
