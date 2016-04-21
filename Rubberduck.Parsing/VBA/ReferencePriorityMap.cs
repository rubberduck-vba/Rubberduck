using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// A <c>Dictionary</c> keyed with a <see cref="VBProject"/>'s ID and valued with an <see cref="int"/> representing a <see cref="Reference"/>'s priority for that project.
    /// </summary>
    public class ReferencePriorityMap : Dictionary<string, int>
    {
        private readonly string _referenceId;

        public ReferencePriorityMap(string referenceId)
        {
            _referenceId = referenceId;
        }

        public string ReferenceId
        {
            get { return _referenceId; }
        }

        public bool IsLoaded { get; set; }

        public override bool Equals(object obj)
        {
            var other = obj as ReferencePriorityMap;
            if (other == null) return false;

            return other.ReferenceId == ReferenceId;
        }

        public override int GetHashCode()
        {
            return _referenceId.GetHashCode();
        }
    }
}