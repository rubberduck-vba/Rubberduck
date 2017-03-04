namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectReference
    {
        private readonly string _referencedProjectId;
        private readonly int _priority;

        public ProjectReference(string referencedProjectId, int priority)
        {
            _priority = priority;
            _referencedProjectId = referencedProjectId;
        }

        public string ReferencedProjectId
        {
            get
            {
                return _referencedProjectId;
            }
        }

        public int Priority
        {
            get
            {
                return _priority;
            }
        }
    }
}
