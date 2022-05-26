namespace Rubberduck.VBEditor
{
    public class CommandBarLocation
    {
        public CommandBarLocation(string parentName, int beforeControlId)
        {
            ParentName = parentName;
            BeforeControlId = beforeControlId;
        }

        public CommandBarLocation(int parentId, int beforeControlId)
        {
            ParentId = parentId;
            BeforeControlId = beforeControlId;
        }

        public string ParentName { get; }
        public int ParentId { get; }
        public int BeforeControlId { get; }
    }
}
