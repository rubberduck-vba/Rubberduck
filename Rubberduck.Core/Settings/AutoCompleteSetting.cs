using System.Xml.Serialization;
using System.Configuration;
using Rubberduck.AutoComplete;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class AutoCompleteSetting : ViewModelBase
    {
        public AutoCompleteSetting() { /* default ctor required for XML serialization */ }

        public AutoCompleteSetting(IAutoComplete autoComplete)
            : this(autoComplete.GetType().Name, autoComplete.IsEnabled) { }

        public AutoCompleteSetting(string key, bool isEnabled)
        {
            Key = key;
            IsEnabled = isEnabled;
        }

        [XmlAttribute]
        public string Key { get; set; }

        private bool _isEnabled;
        [XmlAttribute]
        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        [XmlIgnore]
        public string Description => Resources.Settings.AutoCompletesPage.ResourceManager.GetString(Key + "Description");

        public override bool Equals(object obj)
        {
            var other = obj as AutoCompleteSetting;
            return other != null && other.Key == Key;
        }

        public override int GetHashCode()
        {
            return VBEditor.HashCode.Compute(Key);
        }
    }
}