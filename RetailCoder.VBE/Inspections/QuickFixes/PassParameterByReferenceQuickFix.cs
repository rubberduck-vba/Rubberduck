using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : QuickFixBase
    {
        private Declaration _target;

        public PassParameterByReferenceQuickFix(Declaration target, QualifiedSelection selection)
            : base(target.Context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var argContext = QuickFixHelper.GetArgContexts(Context.Parent.Parent)
                .SingleOrDefault(parameter => Identifier.GetName(parameter.unrestrictedIdentifier())
                    .Equals(_target.IdentifierName));

            module.ReplaceToken(argContext.BYVAL().Symbol,Tokens.ByRef);
        }
    }
}