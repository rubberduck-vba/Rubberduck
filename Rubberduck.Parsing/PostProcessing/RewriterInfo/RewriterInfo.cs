using System;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public struct RewriterInfo : IEquatable<RewriterInfo>
    {
        private readonly string _replacement;
        private readonly int _startTokenIndex;
        private readonly int _stopTokenIndex;

        public RewriterInfo(int startTokenIndex, int stopTokenIndex)
            : this(string.Empty, startTokenIndex, stopTokenIndex) { }

        public RewriterInfo(string replacement, int startTokenIndex, int stopTokenIndex)
        {
            _replacement = replacement;
            _startTokenIndex = startTokenIndex;
            _stopTokenIndex = stopTokenIndex;
        }

        public string Replacement { get { return _replacement; } }
        public int StartTokenIndex { get { return _startTokenIndex; } }
        public int StopTokenIndex { get { return _stopTokenIndex; } }

        public static RewriterInfo None { get { return default(RewriterInfo); } }

        public bool Equals(RewriterInfo other)
        {
            return other.Replacement == Replacement
                   && other.StartTokenIndex == StartTokenIndex
                   && other.StopTokenIndex == StopTokenIndex;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            return Equals((RewriterInfo)obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Replacement, StartTokenIndex, StopTokenIndex);
        }
    }
}