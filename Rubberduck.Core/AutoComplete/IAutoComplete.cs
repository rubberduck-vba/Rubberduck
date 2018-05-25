using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public interface IAutoComplete
    {
        string InputToken { get; }
        string OutputToken { get; }
        bool Execute(AutoCompleteEventArgs e);
        bool IsEnabled { get; }
    }
}
