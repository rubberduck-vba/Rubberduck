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
    public class ReplaceGlobalModifierQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteGlobalInspection)
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
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var selection = result.Context.GetSelection();

            // bug: this should make a test fail somewhere - what if identifier is one of many declarations on a line?
            module.ReplaceLine(selection.StartLine, Tokens.Public + ' ' + result.Context.GetText());
        }

        public string Description(IInspectionResult)
        {
            return InspectionsUI.ObsoleteGlobalInspectionQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}