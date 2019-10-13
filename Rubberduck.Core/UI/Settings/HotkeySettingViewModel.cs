using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class HotkeySettingViewModel : ViewModelBase
    {
        private readonly HotkeySetting wrapped;

        public HotkeySettingViewModel(HotkeySetting wrapped)
        {
            this.wrapped = wrapped;
        }

        public HotkeySetting Unwrap() { return wrapped; }

        public string Key1
        {
            get { return wrapped.Key1; }
            set { wrapped.Key1 = value; OnPropertyChanged(); }
        }

        public bool IsEnabled
        {
            get { return wrapped.IsEnabled; }
            set { wrapped.IsEnabled = value; OnPropertyChanged(); }
        }

        public bool HasShiftModifier
        {
            get { return wrapped.HasShiftModifier; }
            set { wrapped.HasShiftModifier = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsValid)); }
        }

        public bool HasAltModifier
        {
            get { return wrapped.HasAltModifier; }
            set { wrapped.HasAltModifier = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsValid)); }
        }

        public bool HasCtrlModifier
        {
            get { return wrapped.HasCtrlModifier; }
            set { wrapped.HasCtrlModifier = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsValid)); }
        }

        public string CommandTypeName
        {
            get { return wrapped.CommandTypeName; }
            set { wrapped.CommandTypeName = value; OnPropertyChanged(); }
        }

        public bool IsValid { get { return wrapped.IsValid;  } }
        // FIXME If this is the only use, the property should be inlined to here
        public string Prompt { get { return wrapped.Prompt;  } }
    }
}