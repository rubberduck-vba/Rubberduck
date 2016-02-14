using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class IndenterSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public IndenterSettingsViewModel(Configuration config)
        {
            AlignContinuations = config.UserSettings.IndenterSettings.AlignContinuations;
            AlignDim = config.UserSettings.IndenterSettings.AlignDim;
            AlignDimColumn = 15;
            AlignEndOfLine = config.UserSettings.IndenterSettings.AlignEndOfLine;
            AlignIgnoreOps = config.UserSettings.IndenterSettings.AlignIgnoreOps;
        }
        
        #region Properties

        private bool _alignContinuations;
        public bool AlignContinuations
        {
            get { return _alignContinuations; }
            set
            {
                if (_alignContinuations != value)
                {
                    _alignContinuations = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _alignDim;
        public bool AlignDim
        {
            get { return _alignDim; }
            set
            {
                if (_alignDim != value)
                {
                    _alignDim = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _alignDimColumn;
        public int AlignDimColumn
        {
            get { return _alignDimColumn; }
            set
            {
                if (_alignDimColumn != value)
                {
                    _alignDimColumn = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _alignEndOfLine;
        public bool AlignEndOfLine
        {
            get { return _alignEndOfLine; }
            set {
                if (_alignEndOfLine != value)
                {
                    _alignEndOfLine = value;
                    OnPropertyChanged();
                } 
            }
        }

        private bool _alignIgnoreOps;
        public bool AlignIgnoreOps
        {
            get { return _alignIgnoreOps; }
            set
            {
                if (_alignIgnoreOps != value)
                {
                    _alignIgnoreOps = value;
                    OnPropertyChanged();
                }
            }
        }

        #endregion

        public void UpdateConfig(Configuration config) {}
        public void SetToDefaults(Configuration config) {}
    }
}