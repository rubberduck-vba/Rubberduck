using System;
using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.QuickFixes
{
    /// <summary>
    /// An interface that abstracts a runnable code inspection.
    /// </summary>
    public interface IQuickFixProvider
    {
        IEnumerable<IQuickFix> QuickFixes(Type inspectionType);

        IEnumerable<IQuickFix> QuickFixes(IInspectionResult result);

        void Fix(IQuickFix fix, IInspectionResult result);

        void Fix(IQuickFix fix, IEnumerable<IInspectionResult> resultsToFix);

        void FixInProcedure(IQuickFix fix, QualifiedMemberName? selection, Type inspectionType, IEnumerable<IInspectionResult> allResults);

        void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> allResults);

        void FixInProject(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> allResults);

        void FixAll(IQuickFix fix, Type inspectionType, IEnumerable<IInspectionResult> allResults);

        bool HasQuickFixes(IInspectionResult inspectionResult);

        bool CanFix(IQuickFix fix, IInspectionResult result);
    }
}
