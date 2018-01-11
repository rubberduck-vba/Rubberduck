using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using System.Xml.Serialization;
using Rubberduck.Common;

namespace Rubberduck.Settings
{
    public interface IGeneralSettings 
    {
        DisplayLanguageSetting Language { get; set; }
        bool CanShowSplash { get; set; }
        bool CanCheckVersion { get; set; }
        bool IsSmartIndenterPrompted { get; set; }
        bool IsAutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        int MinimumLogLevel { get; set; }
        List<ExperimentalFeatures> EnableExperimentalFeatures { get; set; }
    }

    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings, IEquatable<GeneralSettings>
    {
        public DisplayLanguageSetting Language { get; set; }
        public bool CanShowSplash { get; set; }
        public bool CanCheckVersion { get; set; }
        public bool IsSmartIndenterPrompted { get; set; }
        public bool IsAutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }

        private int _logLevel;
        public int MinimumLogLevel
        {
            get => _logLevel;
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

        public List<ExperimentalFeatures> EnableExperimentalFeatures { get; set; }

        public GeneralSettings()
        {
            Language = new DisplayLanguageSetting("en-US");
            CanShowSplash = true;
            CanCheckVersion = true;
            IsSmartIndenterPrompted = false;
            IsAutoSaveEnabled = false;
            AutoSavePeriod = 10;
            MinimumLogLevel = LogLevel.Off.Ordinal;
            EnableExperimentalFeatures = new List<ExperimentalFeatures>();
        }

        public bool Equals(GeneralSettings other)
        {
            return other != null &&
                   Language.Equals(other.Language) &&
                   CanShowSplash == other.CanShowSplash &&
                   CanCheckVersion == other.CanCheckVersion &&
                   IsSmartIndenterPrompted == other.IsSmartIndenterPrompted &&
                   IsAutoSaveEnabled == other.IsAutoSaveEnabled &&
                   AutoSavePeriod == other.AutoSavePeriod &&
                   MinimumLogLevel == other.MinimumLogLevel &&
                   EnableExperimentalFeatures.All(a => other.EnableExperimentalFeatures.Contains(a)) &&
                   EnableExperimentalFeatures.Count == other.EnableExperimentalFeatures.Count;
        }
    }
}