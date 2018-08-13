namespace Rubberduck.VBEditor
{
    public class CommandBarLocation
    {
        public CommandBarLocation(int parentId, int beforeControlId)
        {
            ParentId = parentId;
            BeforeControlId = beforeControlId;
        }

        public int ParentId { get; }
        public int BeforeControlId { get; }
    }
}
