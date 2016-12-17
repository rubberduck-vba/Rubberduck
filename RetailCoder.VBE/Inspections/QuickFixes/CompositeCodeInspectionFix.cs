using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class CompositeCodeInspectionFix : QuickFixBase
    {
        private readonly QuickFixBase _root;
        private readonly List<QuickFixBase> _children;

        public CompositeCodeInspectionFix(QuickFixBase root)
            : base(root.Context, root.Selection, root.Description)
        {
            _root = root;
            _children = new List<QuickFixBase>();
        }

        public void AddChild(QuickFixBase quickFix)
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
