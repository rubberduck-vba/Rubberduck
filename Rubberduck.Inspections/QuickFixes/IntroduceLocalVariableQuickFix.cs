using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IntroduceLocalVariableQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(UndeclaredVariableInspection)
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

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;

        public void Fix(IInspectionResult result)
        {
            var instruction = Tokens.Dim + ' ' + result.Target.IdentifierName + ' ' + Tokens.As + ' ' + Tokens.Variant;
            result.QualifiedSelection.QualifiedName.Component.CodeModule.InsertLines(result.QualifiedSelection.Selection.StartLine, instruction);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.IntroduceLocalVariableQuickFix;
        }
    }
}