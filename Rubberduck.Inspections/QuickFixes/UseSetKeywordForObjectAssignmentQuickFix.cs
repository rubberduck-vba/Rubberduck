using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class UseSetKeywordForObjectAssignmentQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObjectVariableNotSetInspection)
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
            var codeLine = module.GetLines(result.QualifiedSelection.Selection.StartLine, 1);

            var letStatementLeftSide = result.Context.GetText();
            var setStatementLeftSide = Tokens.Set + ' ' + letStatementLeftSide;

            var correctLine = codeLine.Replace(letStatementLeftSide, setStatementLeftSide);
            module.ReplaceLine(result.QualifiedSelection.Selection.StartLine, correctLine);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SetObjectVariableQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}