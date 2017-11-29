namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectReference
    {
        public ProjectReference(string referencedProjectId, int priority)
        {
            Priority = priority;
            ReferencedProjectId = referencedProjectId;
        }

        public string ReferencedProjectId { get; }

        public int Priority { get; }
    }
}
