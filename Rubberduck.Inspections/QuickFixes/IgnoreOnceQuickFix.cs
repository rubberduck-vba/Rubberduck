using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IgnoreOnceQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>();

        public IgnoreOnceQuickFix(IEnumerable<IInspection> inspections)
        {
            _supportedInspections.UnionWith(inspections.Select(i => i.GetType()));
        }

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public static void RemoveSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Remove(inspectionType);
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;

        public void Fix(IInspectionResult result)
        {
            var annotationText = "'" + Annotations.AnnotationMarker + Annotations.IgnoreInspection + ' ' + result.Inspection.AnnotationName;

            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var insertLine = result.QualifiedSelection.Selection.StartLine;
            while (insertLine != 1 && module.GetLines(insertLine - 1, 1).EndsWith(" _"))
            {
                insertLine--;
            }
            var codeLine = insertLine == 1 ? string.Empty : module.GetLines(insertLine - 1, 1);
            var ignoreAnnotation = "'" + Annotations.AnnotationMarker + Annotations.IgnoreInspection;

            int commentStart;
            if (codeLine.HasComment(out commentStart) && codeLine.Substring(commentStart).StartsWith(ignoreAnnotation))
            {
                var indentation = codeLine.Length - codeLine.TrimStart().Length;
                annotationText = $"{new string(' ', indentation)}{annotationText},{codeLine.Substring(indentation + ignoreAnnotation.Length)}";
                module.ReplaceLine(insertLine - 1, annotationText);
            }
            else
            {
                module.InsertLines(insertLine, annotationText);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.IgnoreOnce;
        }
    }
}
