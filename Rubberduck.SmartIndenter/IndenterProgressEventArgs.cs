using System;

namespace Rubberduck.SmartIndenter
{
    public class IndenterProgressEventArgs : EventArgs
    {
        public IndenterProgressEventArgs(string componentName, int progress, int max)
        {
            Progress = progress;
            Max = max;
            ComponentName = componentName;
        }

        public int Progress { get; }

        public string ComponentName { get; }

        public int Max { get; }
    }
}
