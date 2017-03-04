using System;
using NLog;
using System.Xml.Serialization;
using Rubberduck.Common;

namespace Rubberduck.Settings
{
    public interface IGeneralSettings 
    {
        DisplayLanguageSetting Language { get; set; }
        bool ShowSplash { get; set; }
        bool CheckVersion { get; set; }
        bool SmartIndenterPrompted { get; set; }
        bool AutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        //char Delimiter { get; set; }
        int MinimumLogLevel { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings, IEquatable<GeneralSettings>
    {
        public DisplayLanguageSetting Language { get; set; }
        public bool ShowSplash { get; set; }
        public bool CheckVersion { get; set; }
        public bool SmartIndenterPrompted { get; set; }
        public bool AutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }
        //public char Delimiter { get; set; }

        private int _logLevel;
        public int MinimumLogLevel
        {
            get { return _logLevel; }
            set
            {
                if (value < LogLevelHelper.MinLogLevel())
                {
                    _logLevel = LogLevelHelper.MinLogLevel();
                }
                else if (value > LogLevelHelper.MaxLogLevel())
                {
                    _logLevel = LogLevelHelper.MaxLogLevel();
                }
                else
                {
                    _logLevel = value;
                }               
            }
        }

        public GeneralSettings()
        {
            Language = new DisplayLanguageSetting("en-US");
            ShowSplash = true;
            CheckVersion = true;
            SmartIndenterPrompted = false;
            AutoSaveEnabled = false;
            AutoSavePeriod = 10;
            //Delimiter = '.';
            MinimumLogLevel = LogLevel.Off.Ordinal;
        }

        public bool Equals(GeneralSettings other)
        {
            return other != null &&
                   Language.Equals(other.Language) &&
                   ShowSplash == other.ShowSplash &&
                   CheckVersion == other.CheckVersion &&
                   SmartIndenterPrompted == other.SmartIndenterPrompted &&
                   AutoSaveEnabled == other.AutoSaveEnabled &&
                   AutoSavePeriod == other.AutoSavePeriod &&
                   //Delimiter.Equals(other.Delimiter) &&
                   MinimumLogLevel == other.MinimumLogLevel;
        }
    }
}