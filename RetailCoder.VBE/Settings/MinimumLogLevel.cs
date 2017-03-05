using System.Globalization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    public sealed class MinimumLogLevel
    {
        private readonly int _ordinal;
        private readonly string _name;

        public MinimumLogLevel(int ordinal, string logLevelName)
        {
            _ordinal = ordinal;
            _name = RubberduckUI.ResourceManager.GetString("GeneralSettings_" + logLevelName + "LogLevel", CultureInfo.CurrentUICulture);
        }

        public int Ordinal
        {
            get
            {
                return _ordinal;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }
    }
}
