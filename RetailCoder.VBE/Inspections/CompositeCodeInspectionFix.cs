using Antlr4.Runtime;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
