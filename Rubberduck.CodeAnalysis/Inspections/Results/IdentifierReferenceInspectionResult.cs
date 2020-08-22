using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Results
{
    internal class IdentifierReferenceInspectionResult : InspectionResultBase
    {
        public IdentifierReference Reference { get; }

        public IdentifierReferenceInspectionResult(
            IInspection inspection, 
            string description, 
            DeclarationFinder finder, 
            IdentifierReference reference,
            ICollection<string> disabledQuickFixes = null) 
            : base(inspection,
                 description,
                 reference.QualifiedModuleName,
                 reference.Context,
                 reference.Declaration,
                 new QualifiedSelection(reference.QualifiedModuleName, reference.Context.GetSelection()),
                 GetQualifiedMemberName(finder, reference),
                 disabledQuickFixes)
        {
            Reference = reference;
        }

        private static QualifiedMemberName? GetQualifiedMemberName(DeclarationFinder finder, IdentifierReference reference)
        {
            var members = finder.Members(reference.QualifiedModuleName);
            return members.SingleOrDefault(m => reference.Context.IsDescendentOf(m.Context))?.QualifiedName;
        }

        public override bool ChangesInvalidateResult(ICollection<QualifiedModuleName> modifiedModules)
        {
            return Target != null && modifiedModules.Contains(Target.QualifiedModuleName)
                   || base.ChangesInvalidateResult(modifiedModules);
        }
    }

    internal class IdentifierReferenceInspectionResult<T> : IdentifierReferenceInspectionResult, IWithInspectionResultProperties<T>
    {
        public IdentifierReferenceInspectionResult(
            IInspection inspection, 
            string description, 
            DeclarationFinder finder, 
            IdentifierReference reference, 
            T properties,
            ICollection<string> disabledQuickFixes = null) 
            : base(
                inspection,
                description,
                finder, 
                reference,
                disabledQuickFixes)
        {
            Properties = properties;
        }

        public T Properties { get; }
    }
}
