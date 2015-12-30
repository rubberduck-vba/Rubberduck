using System;

namespace Rubberduck.AutoSave
{
    public interface IAutoSaveSettings
    {
        bool IsEnabled { get; set; }
        double TimerDelay { get; set; }

        event EventHandler IsEnabledChanged;
        event EventHandler TimerDelayChanged;
    }

    public class AutoSaveSettings : IAutoSaveSettings
    {
        public event EventHandler IsEnabledChanged;
        public event EventHandler TimerDelayChanged;

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
                    OnIsEnabledChanged();
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
                    OnTimerDelayChanged();
                }
            }
        }

        protected virtual void OnIsEnabledChanged()
        {
            var handler = IsEnabledChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        protected virtual void OnTimerDelayChanged()
        {
            var handler = TimerDelayChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
