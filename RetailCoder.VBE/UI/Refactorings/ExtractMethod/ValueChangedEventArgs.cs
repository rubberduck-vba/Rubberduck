using System;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ValueChangedEventArgs<TValue> : EventArgs
    {
        public ValueChangedEventArgs(TValue newValue)
        {
            _newValue = newValue;
        }

        private readonly TValue _newValue;
        public TValue NewValue { get { return _newValue; } }
    }
}