using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Rubberduck.AutoSave
{
    public interface IAutoSaveSettings : INotifyPropertyChanged
    {
        bool IsEnabled { get; set; }
        double TimerDelay { get; set; }
    }

    public class AutoSaveSettings : IAutoSaveSettings
    {
        public AutoSaveSettings(bool isEnabled = true, int timerDelay = 600000)
        {
            IsEnabled = isEnabled;
            TimerDelay = timerDelay;
        }

        private bool _isEnabled;
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

        private double _timerDelay;
        public double TimerDelay
        {
            get { return _timerDelay; }
            set
            {
                if (Math.Abs(_timerDelay - value) > .1)
                {
                    _timerDelay = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
