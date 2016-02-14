using System.Collections.Generic;

namespace Rubberduck.Inspections
{
    public sealed class CompositeCodeInspectionFix : CodeInspectionQuickFix
    {
        private readonly CodeInspectionQuickFix _root;
        private readonly List<CodeInspectionQuickFix> _children;

        public CompositeCodeInspectionFix(CodeInspectionQuickFix root)
            : base(root.Context, root.Selection, root.Description)
        {
            _root = root;
            _children = new List<CodeInspectionQuickFix>();
        }

        public void AddChild(CodeInspectionQuickFix quickFix)
        {
            _children.Add(quickFix);
        }

        public override void Fix()
        {
            _root.Fix();
            _children.ForEach(child => child.Fix());
        }
    }
}
