using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.Common;
using Rubberduck.Resources;

namespace Rubberduck.Settings
{
    public interface IGeneralSettings 
    {
        DisplayLanguageSetting Language { get; set; }
        bool CanShowSplash { get; set; }
        bool CanCheckVersion { get; set; }
        bool IncludePreRelease { get; set; }
        bool CompileBeforeParse { get; set; }
        bool IsSmartIndenterPrompted { get; set; }
        bool IsAutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        bool UserEditedLogLevel { get; set; }
        int MinimumLogLevel { get; set; }
        bool SetDpiUnaware { get; set; }
        bool EnableFolderDragAndDrop { get; set; }
        List<ExperimentalFeature> EnableExperimentalFeatures { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings, IEquatable<GeneralSettings>
    {
        private DisplayLanguageSetting _language = new DisplayLanguageSetting(Locales.DefaultCulture.Name);

        public DisplayLanguageSetting Language
        {
            get => _language;
            set
            {
                if (Locales.AvailableCultures.Exists(culture => culture.Name.Equals(value.Code, StringComparison.OrdinalIgnoreCase)))
                {
                    _language = value;
                }
            }
        }

        public bool CanShowSplash { get; set; }
        public bool CanCheckVersion { get; set; }
        public bool IncludePreRelease { get; set; }
        public bool CompileBeforeParse { get; set; }
        public bool IsSmartIndenterPrompted { get; set; }
        public bool IsAutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }

        public bool UserEditedLogLevel { get; set; }

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

        public bool SetDpiUnaware { get; set; }
        public bool EnableFolderDragAndDrop { get; set; }

        public List<ExperimentalFeature> EnableExperimentalFeatures { get; set; } = new List<ExperimentalFeature>();

        public GeneralSettings()
        {
            //Enforce non-default default value for members
            //In other words, if we want a bool to default to
            //true, it must be set here for correct behavior
            CompileBeforeParse = true;
        }

        public bool Equals(GeneralSettings other)
        {
            return other != null &&
                   Language.Equals(other.Language) &&
                   CanShowSplash == other.CanShowSplash &&
                   CanCheckVersion == other.CanCheckVersion &&
                   IncludePreRelease == other.IncludePreRelease &&
                   CompileBeforeParse == other.CompileBeforeParse &&
                   IsSmartIndenterPrompted == other.IsSmartIndenterPrompted &&
                   IsAutoSaveEnabled == other.IsAutoSaveEnabled &&
                   AutoSavePeriod == other.AutoSavePeriod &&
                   UserEditedLogLevel == other.UserEditedLogLevel &&
                   MinimumLogLevel == other.MinimumLogLevel &&                   
                   EnableExperimentalFeatures.Count == other.EnableExperimentalFeatures.Count &&
                   EnableExperimentalFeatures.All(other.EnableExperimentalFeatures.Contains) &&
                   SetDpiUnaware == other.SetDpiUnaware &&
                   EnableFolderDragAndDrop == other.EnableFolderDragAndDrop;
        }
    }
}