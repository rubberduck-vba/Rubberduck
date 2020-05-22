using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.UI.Inspections
{
    public static class InspectionResultComparer
    {
        public static Comparer<IInspectionResult> InspectionType { get; } = new CompareByInspectionType();

        public static Comparer<IInspectionResult> Name { get; } = new CompareByInspectionName();

        public static Comparer<IInspectionResult> Location { get; } = new CompareByLocation();

        public static Comparer<IInspectionResult> Severity { get; } = new CompareBySeverity();
    }

    public class CompareByInspectionType : Comparer<IInspectionResult>
    {
        public override int Compare(IInspectionResult x, IInspectionResult y)
        {
            if (x == y)
            {
                return 0;
            }

            if (x?.Inspection?.InspectionType is null)
            {
                return -1;
            }

            return x.Inspection.InspectionType.CompareTo(y?.Inspection?.InspectionType);
        }
    }

    public class CompareByInspectionName : Comparer<IInspectionResult>
    {
        public override int Compare(IInspectionResult x, IInspectionResult y)
        {
            if (x == y)
            {
                return 0;
            }

            if (x?.Inspection?.Name is null)
            {
                return -1;
            }

            return x.Inspection.Name.CompareTo(y?.Inspection?.Name);
        }
    }

    public class CompareByLocation : Comparer<IInspectionResult>
    {
        public override int Compare(IInspectionResult x, IInspectionResult y)
        {
            if (x == y)
            {
                return 0;
            }

            var first = x?.QualifiedSelection.QualifiedName;

            if (first is null)
            {
                return -1;
            }

            var second = y?.QualifiedSelection.QualifiedName;

            if (second is null)
            {
                return 1;
            }

            var nameComparison = first.Value.Name.CompareTo(second.Value.Name);
            if (nameComparison != 0)
            {
                return nameComparison;
            }

            return x.QualifiedSelection.CompareTo(y.QualifiedSelection);
        }
    }

    public class CompareBySeverity : Comparer<IInspectionResult>
    {
        public override int Compare(IInspectionResult x, IInspectionResult y)
        {
            if (x == y)
            {
                return 0;
            }

            if (y?.Inspection is null)
            {
                return 1;
            }

            // The CodeInspectionSeverity enum is ordered least severe to most. We want the opposite. 
            return y.Inspection.Severity.CompareTo(x?.Inspection?.Severity);
        }
    }
}
