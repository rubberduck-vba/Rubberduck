using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IProviderPresenter
    {
        ISourceControlProvider Provider { get; set; }
    }
}
