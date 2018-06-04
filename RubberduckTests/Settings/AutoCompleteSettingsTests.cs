using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class AutoCompleteSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var autoCompleteSettings = new AutoCompleteSettings
            {
                AutoCompletes = new HashSet<AutoCompleteSetting>(new[]
                {
                    new AutoCompleteSetting("AutoCompleteClosingString", true),
                    new AutoCompleteSetting("SomeDisabledAutoComplete", false)
                })
            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNonDefaultConfig()
        {
            var autoCompleteSettings = new AutoCompleteSettings
            {
                AutoCompletes = new HashSet<AutoCompleteSetting>(new[]
                {
                    new AutoCompleteSetting("AutoCompleteClosingString", false),
                    new AutoCompleteSetting("SomeDisabledAutoComplete", true)
                })
            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        // todo: test settings viewmodel here
    }
}
