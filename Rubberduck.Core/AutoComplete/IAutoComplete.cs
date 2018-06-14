using Rubberduck.Settings;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public interface IAutoComplete
    {
        string InputToken { get; }
        string OutputToken { get; }
        bool IsMatch(string code);
        bool Execute(AutoCompleteEventArgs e, AutoCompleteSettings settings);
        bool IsInlineCharCompletion { get; }
        bool IsEnabled { get; set; }
    }
}
