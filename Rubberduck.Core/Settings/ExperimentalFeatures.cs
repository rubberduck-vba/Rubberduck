using System.Xml.Serialization;
using Rubberduck.UI;
using Rubberduck.Resources.Experimentals;

namespace Rubberduck.Settings
{
    public class ExperimentalFeature : ViewModelBase
    {
        private bool _isEnabled;
        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                _isEnabled = value;
                OnPropertyChanged();
            }
        }

        private string _key;
        public string Key
        {
            get { return _key; }
            set
            {
                _key = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(DisplayValue));
            }
        }

        [XmlIgnore]
        public string DisplayValue => Key == null ? string.Empty : ExperimentalNames.ResourceManager.GetString(Key);

        public override string ToString()
        {
            return Key;
        }

        public override bool Equals(object obj)
        {
            return obj is ExperimentalFeature value &&
                   value.IsEnabled == IsEnabled &&
                   value.Key == Key;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (IsEnabled.GetHashCode() * 397) ^ (Key != null ? Key.GetHashCode() : 0);
            }
        }
    }
}