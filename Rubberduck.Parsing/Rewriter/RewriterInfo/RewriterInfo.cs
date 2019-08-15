using System;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public struct RewriterInfo : IEquatable<RewriterInfo>
    {
        public RewriterInfo(int startTokenIndex, int stopTokenIndex)
            : this(string.Empty, startTokenIndex, stopTokenIndex) { }

        public RewriterInfo(string replacement, int startTokenIndex, int stopTokenIndex)
        {
            Replacement = replacement;
            StartTokenIndex = startTokenIndex;
            StopTokenIndex = stopTokenIndex;
        }

        public string Replacement { get; }
        public int StartTokenIndex { get; }
        public int StopTokenIndex { get; }

        public static RewriterInfo None => default;

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