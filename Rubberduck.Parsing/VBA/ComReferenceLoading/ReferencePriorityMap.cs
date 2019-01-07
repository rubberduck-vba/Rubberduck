using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    /// <summary>
    /// A <c>Dictionary</c> keyed with a project's ID and valued with an int representing a reference's priority for that project.
    /// </summary>
    public class ReferencePriorityMap : Dictionary<string, int>
    {
        private readonly string _referencedProjectId;

        public ReferencePriorityMap(string referencedProjectId)
        {
            _referencedProjectId = referencedProjectId;
        }

        public string ReferencedProjectId => _referencedProjectId;

        public bool IsLoaded { get; set; }

        public override bool Equals(object obj)
        {
            if (!(obj is ReferencePriorityMap other))
            {
                return false;
            }

            return other.ReferencedProjectId == ReferencedProjectId;
        }

        public override int GetHashCode()
        {
            return _referencedProjectId.GetHashCode();
        }
    }
}
