using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IControlViewModel
    {
        ISourceControlProvider Provider { get; set; }
    }
}