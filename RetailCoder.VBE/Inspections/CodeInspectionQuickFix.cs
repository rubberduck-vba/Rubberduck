namespace Rubberduck.Inspections
{
    public abstract class CodeInspectionQuickFix
    {
        private readonly string _description;

        public CodeInspectionQuickFix(string description)
        {
            _description = description;
        }

        public string Description { get { return _description; } }

        public abstract void Fix();
    }
}