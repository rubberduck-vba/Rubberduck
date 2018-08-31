using Rubberduck.AutoComplete;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IAutoCompleteSettings
    {
        HashSet<AutoCompleteSetting> AutoCompletes { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class AutoCompleteSettings : IAutoCompleteSettings, IEquatable<AutoCompleteSettings>
    {
        [XmlArrayItem("AutoComplete", IsNullable = false)]
        public HashSet<AutoCompleteSetting> AutoCompletes { get; set; }

        [XmlAttribute]
        public bool IsEnabled { get; set; }

        [XmlAttribute]
        public bool CompleteBlockOnTab { get; set; }
        
        [XmlAttribute]
        public bool CompleteBlockOnEnter { get; set; }

        [XmlAttribute]
        public bool EnableSmartConcat { get; set; }

        public AutoCompleteSettings() : this(Enumerable.Empty<AutoCompleteSetting>())
        {
            /* default constructor required for XML serialization */
        }

        public AutoCompleteSettings(IEnumerable<AutoCompleteSetting> defaultSettings)
        {
            AutoCompletes = new HashSet<AutoCompleteSetting>(defaultSettings);
        }

        public AutoCompleteSetting GetSetting<TAutoComplete>() where TAutoComplete : IAutoComplete
        {
            return AutoCompletes.FirstOrDefault(s => typeof(TAutoComplete).Name.Equals(s.Key))
                ?? GetSetting(typeof(TAutoComplete));
        }

        public AutoCompleteSetting GetSetting(Type autoCompleteType)
        {
            try
            {
                var existing = AutoCompletes.FirstOrDefault(s => autoCompleteType.Name.Equals(s.Key));
                if (existing != null)
                {
                    return existing;
                }
                var proto = Convert.ChangeType(Activator.CreateInstance(autoCompleteType, new object[] { null }), autoCompleteType);
                var setting = new AutoCompleteSetting(proto as IAutoComplete);
                AutoCompletes.Add(setting);
                return setting;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool Equals(AutoCompleteSettings other)
        {
            return other != null && AutoCompletes.SequenceEqual(other.AutoCompletes);
        }
    }
}
