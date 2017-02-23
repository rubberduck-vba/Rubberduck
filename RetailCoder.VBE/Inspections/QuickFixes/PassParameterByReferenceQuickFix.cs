using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : QuickFixBase
    {
        private Declaration _target;
        private QuickFixHelper _quickFixHelper;

        public PassParameterByReferenceQuickFix(Declaration target, QualifiedSelection selection)
            : base(target.Context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
            _target = target;
            _quickFixHelper = new QuickFixHelper(target, selection);
        }

        public override void Fix()
        {
            var argCtxt = GetArgContextForIdentifier(Context.Parent.Parent, _target.IdentifierName);

            _quickFixHelper.ReplaceTerminalNodeTextInCodeModule(argCtxt.BYVAL(), Tokens.ByRef);
        }
        private VBAParser.ArgContext GetArgContextForIdentifier(RuleContext context, string identifier)
        {
            var args = _quickFixHelper.GetArgContextsForContext(context);
            return args.SingleOrDefault(parameter =>
                    Identifier.GetName(parameter.unrestrictedIdentifier()).Equals(identifier));
        }
    }
}