using System;

namespace Rubberduck.SmartIndenter
{
    public class IndenterProgressEventArgs : EventArgs
    {
        private readonly int _progress;
        private readonly int _max;
        private readonly string _componentName;

        public IndenterProgressEventArgs(string componentName, int progress, int max)
        {
            _progress = progress;
            _max = max;
            _componentName = componentName;
        }

        public int Progress
        {
            get { return _progress; }
        }

        public string ComponentName
        {
            get { return _componentName; }
        }

        public int Max
        {
            get { return _max; }
        }
    }
}
