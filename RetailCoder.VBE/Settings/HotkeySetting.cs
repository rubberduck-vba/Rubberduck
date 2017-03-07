using System.Globalization;
using System.Xml.Serialization;
using Rubberduck.UI;

namespace Rubberduck.Settings
{
    public class HotkeySetting
    {
        public const string KeyModifierAlt = "%";
        public const string KeyModifierCtrl = "^";
        public const string KeyModifierShift = "+";

        public string Name { get; set; }
        public string Key1 { get; set; }
        /// <summary>
        /// For 2-step hotkeys, the 2nd key to press. Note: hidden until 2-step hotkeys are an actual thing.
        /// </summary>
        public string Key2 { get; set; }
        public bool IsEnabled { get; set; }
        public bool HasShiftModifier { get; set; }
        public bool HasAltModifier { get; set; }
        public bool HasCtrlModifier { get; set; }

        [XmlIgnore]
        public string Prompt
        {
            get { return RubberduckUI.ResourceManager.GetString("HotkeyDescription_" + Name, CultureInfo.CurrentUICulture); } 
        }

        public override string ToString()
        {
            return string.Format("{0}{1}{2}{3}",
                HasCtrlModifier ? KeyModifierCtrl : string.Empty,
                HasShiftModifier ? KeyModifierShift : string.Empty,
                HasAltModifier ? KeyModifierAlt : string.Empty,
                Key1);
        }

        public override bool Equals(object obj)
        {
            var hotkey = obj as HotkeySetting;

            return hotkey != null &&
                   hotkey.Name == Name &&
                   hotkey.Key1 == Key1 &&
                   hotkey.Key2 == Key2 &&
                   hotkey.HasAltModifier == HasAltModifier &&
                   hotkey.HasCtrlModifier == HasCtrlModifier &&
                   hotkey.HasShiftModifier == HasShiftModifier &&
                   hotkey.IsEnabled == IsEnabled;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Name != null ? Name.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Key1 != null ? Key1.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Key2 != null ? Key2.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ HasShiftModifier.GetHashCode();
                hashCode = (hashCode * 397) ^ HasCtrlModifier.GetHashCode();
                hashCode = (hashCode * 397) ^ HasAltModifier.GetHashCode();
                hashCode = (hashCode * 397) ^ IsEnabled.GetHashCode();
                return hashCode;
            }
        }
    }
}