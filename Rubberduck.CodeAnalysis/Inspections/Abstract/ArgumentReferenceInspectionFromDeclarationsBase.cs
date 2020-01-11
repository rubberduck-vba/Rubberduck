using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Inspections.Abstract
{
    public abstract class ArgumentReferenceInspectionFromDeclarationsBase : IdentifierReferenceInspectionFromDeclarationsBase
    {
        protected ArgumentReferenceInspectionFromDeclarationsBase(RubberduckParserState state) 
            : base(state) { }

        protected abstract bool IsUnsuitableArgument(ArgumentReference reference, DeclarationFinder finder);

        protected virtual (bool isResult, object properties) IsUnsuitableArgumentWithAdditionalProperties(ArgumentReference reference, DeclarationFinder finder)
        {
            return (IsUnsuitableArgument(reference, finder), null);
        }

        protected override IEnumerable<(IdentifierReference reference, object properties)> ObjectionableReferences(DeclarationFinder finder)
        {
            return ObjectionableDeclarations(finder)
                .OfType<ModuleBodyElementDeclaration>()
                .SelectMany(declaration => declaration.Parameters)
                .SelectMany(parameter => parameter.ArgumentReferences)
                .Select(reference => (reference, IsResultReferenceWithAdditionalProperties(reference, finder)))
                .Where(tpl => tpl.Item2.isResult)
                .Select(tpl => ((IdentifierReference) tpl.reference, tpl.Item2.properties)); ;
        }

        protected override (bool isResult, object properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            if (!(reference is ArgumentReference argumentReference))
            {
                return (false, null);
            }

            return IsUnsuitableArgumentWithAdditionalProperties(argumentReference, finder);
        }
    }
}
