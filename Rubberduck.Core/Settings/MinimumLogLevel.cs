using System.Globalization;
using Rubberduck.UI.Settings;

namespace Rubberduck.Settings
{
    public sealed class MinimumLogLevel
    {
        public MinimumLogLevel(int ordinal, string logLevelName)
        {
            Ordinal = ordinal;
            Name = GeneralSettingsUI.ResourceManager.GetString(logLevelName + "LogLevel", CultureInfo.CurrentUICulture);
        }

        public int Ordinal { get; }

        public string Name { get; }
    }
}
