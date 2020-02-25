using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierReferenceInspectionResult : InspectionResultBase
    {
        public IdentifierReference Reference { get; }

        public IdentifierReferenceInspectionResult(
            IInspection inspection, 
            string description, 
            IDeclarationFinderProvider declarationFinderProvider, 
            IdentifierReference reference,
            ICollection<string> disabledQuickFixes = null) 
            : base(inspection,
                 description,
                 reference.QualifiedModuleName,
                 reference.Context,
                 reference.Declaration,
                 new QualifiedSelection(reference.QualifiedModuleName, reference.Context.GetSelection()),
                 GetQualifiedMemberName(declarationFinderProvider, reference),
                 disabledQuickFixes)
        {
            Reference = reference;
        }

        private static QualifiedMemberName? GetQualifiedMemberName(IDeclarationFinderProvider declarationFinderProvider, IdentifierReference reference)
        {
            var members = declarationFinderProvider.DeclarationFinder.Members(reference.QualifiedModuleName);
            return members.SingleOrDefault(m => reference.Context.IsDescendentOf(m.Context))?.QualifiedName;
        }

        public override bool ChangesInvalidateResult(ICollection<QualifiedModuleName> modifiedModules)
        {
            return Target != null && modifiedModules.Contains(Target.QualifiedModuleName)
                   || base.ChangesInvalidateResult(modifiedModules);
        }
    }

    public class IdentifierReferenceInspectionResult<T> : IdentifierReferenceInspectionResult, IWithInspectionResultProperties<T>
    {
        public IdentifierReferenceInspectionResult(
            IInspection inspection, 
            string description, 
            IDeclarationFinderProvider declarationFinderProvider, 
            IdentifierReference reference, 
            T properties,
            ICollection<string> disabledQuickFixes = null) 
            : base(
                inspection,
                description,
                declarationFinderProvider, 
                reference,
                disabledQuickFixes)
        {
            Properties = properties;
        }

        public T Properties { get; }
    }
}
