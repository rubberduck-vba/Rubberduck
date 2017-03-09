using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveUnassignedIdentifierQuickFix : QuickFixBase
    {
        private readonly Declaration _target;

        public RemoveUnassignedIdentifierQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target)
            : base(context, selection, InspectionsUI.RemoveUnassignedIdentifierQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.Remove(_target);
        }
    }
}