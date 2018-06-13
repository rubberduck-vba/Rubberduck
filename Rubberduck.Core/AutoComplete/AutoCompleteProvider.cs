using System.Collections.Generic;
using System.Linq;
using Rubberduck.Settings;
using NLog;

namespace Rubberduck.AutoComplete
{
    public interface IAutoCompleteProvider
    {
        IEnumerable<IAutoComplete> AutoCompletes { get; }
    }

    public class AutoCompleteProvider : IAutoCompleteProvider
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();
        
        public AutoCompleteProvider(IEnumerable<IAutoComplete> autoCompletes)
        {
            var defaults = new DefaultSettings<AutoCompleteSettings>().Default;
            var defaultKeys = defaults.AutoCompletes.Select(x => x.Key);
            var defaultAutoCompletes = autoCompletes.Where(autoComplete => defaultKeys.Contains(autoComplete.GetType().Name));

            foreach (var autoComplete in defaultAutoCompletes)
            {
                autoComplete.IsEnabled = defaults.AutoCompletes.First(setting => setting.Key == autoComplete.GetType().Name).IsEnabled;
            }

            AutoCompletes = autoCompletes;
            _logger.Trace($"{AutoCompletes.Count()} IAutoComplete implementations registered.");
        }

        public IEnumerable<IAutoComplete> AutoCompletes { get; }
    }
}
