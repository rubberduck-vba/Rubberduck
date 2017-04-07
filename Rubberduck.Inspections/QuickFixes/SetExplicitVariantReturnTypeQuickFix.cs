using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SetExplicitVariantReturnTypeQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ImplicitVariantReturnTypeInspection)
        };

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            var procedure = result.Context.GetText();
            // todo: verify that this isn't a bug / test with a procedure that contains parentheses in the body.
            var indexOfLastClosingParen = procedure.LastIndexOf(')');

            var newContent = indexOfLastClosingParen == procedure.Length
                ? procedure + ' ' + Tokens.As + ' ' + Tokens.Variant
                : procedure.Insert(procedure.LastIndexOf(')') + 1, ' ' + Tokens.As + ' ' + Tokens.Variant);
            
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var selection = result.Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SetExplicitVariantReturnTypeQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}