using System.Globalization;
using Rubberduck.Resources;

namespace Rubberduck.Settings
{
    public sealed class MinimumLogLevel
    {
        public MinimumLogLevel(int ordinal, string logLevelName)
        {
            Ordinal = ordinal;
            Name = RubberduckUI.ResourceManager.GetString("GeneralSettings_" + logLevelName + "LogLevel", CultureInfo.CurrentUICulture);
        }

        public int Ordinal { get; }

        public string Name { get; }
    }
}
