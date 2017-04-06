using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RepaceEmptyStringLiteralStatementQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(EmptyStringLiteralInspection)
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

            var literal = (VBAParser.LiteralExpressionContext)result.Context;
            var newCodeLines = module.GetLines(literal.Start.Line, 1).Replace("\"\"", "vbNullString");

            module.ReplaceLine(literal.Start.Line, newCodeLines);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.EmptyStringLiteralInspectionQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}