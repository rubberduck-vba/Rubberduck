using System.Collections.Generic;
using System.Linq;
using Rubberduck.Settings;

namespace Rubberduck.AutoComplete
{
    public interface IAutoCompleteProvider
    {
        IEnumerable<IAutoComplete> AutoCompletes { get; }
    }

    public class AutoCompleteProvider : IAutoCompleteProvider
    {
        public AutoCompleteProvider(IEnumerable<IAutoComplete> autoCompletes)
        {
            var defaults = new DefaultSettings<AutoCompleteSettings>().Default;
            var defaultKeys = defaults.Settings.Select(x => x.Key);
            var defaultAutoCompletes = autoCompletes.Where(autoComplete => defaultKeys.Contains(autoComplete.GetType().Name));

            foreach (var autoComplete in defaultAutoCompletes)
            {
                autoComplete.IsEnabled = defaults.Settings.First(setting => setting.Key == autoComplete.GetType().Name).IsEnabled;
            }

            AutoCompletes = autoCompletes;
        }

        public IEnumerable<IAutoComplete> AutoCompletes { get; }
    }
}
