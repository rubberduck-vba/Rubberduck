using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveCommentQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteCommentSyntaxInspection)
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

            if (module.IsWrappingNullReference)
            {
                return;                
            }

            var start = result.QualifiedSelection.Selection.StartLine;
            var commentLine = module.GetLines(start, result.QualifiedSelection.Selection.LineCount);
            var newLine = commentLine.Substring(0, result.QualifiedSelection.Selection.StartColumn - 1).TrimEnd();

            module.DeleteLines(start, result.QualifiedSelection.Selection.LineCount);
            if (newLine.TrimStart().Length > 0)
            {
                module.InsertLines(start, newLine);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}