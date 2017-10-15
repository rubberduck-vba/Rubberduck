using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : QuickFixBase
    {
        private readonly ICodeModule _codeModule;
        private readonly VBAParser.ArgContext _argContext;

        public PassParameterByReferenceQuickFix(Declaration target, QualifiedSelection selection)
            : base(target.Context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
            _argContext = target.Context as VBAParser.ArgContext;
            _codeModule = Selection.QualifiedName.Component.CodeModule;
        }

        public override void Fix()
        {
            _codeModule.ReplaceToken(_argContext.BYVAL().Symbol, Tokens.ByRef);
        }
    }
}