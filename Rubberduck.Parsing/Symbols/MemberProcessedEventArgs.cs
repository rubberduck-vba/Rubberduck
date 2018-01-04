using System;

namespace Rubberduck.Parsing.Symbols
{
    public class MemberProcessedEventArgs : EventArgs
    {
        public MemberProcessedEventArgs(string name)
        {
            MemberName = name;
        }

        public string MemberName { get; }
    }
}
