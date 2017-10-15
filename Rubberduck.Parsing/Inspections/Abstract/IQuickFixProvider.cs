using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    /// <summary>
    /// An interface that abstracts a runnable code inspection.
    /// </summary>
    public interface IQuickFixProvider
    {
        IEnumerable<IQuickFix> QuickFixes(IInspectionResult result);

        void Fix(IQuickFix fix, IInspectionResult result);

        void FixInProcedure(IQuickFix fix, QualifiedMemberName? selection, Type inspectionType,
            IEnumerable<IInspectionResult> results);

        void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType,
            IEnumerable<IInspectionResult> results);

        void FixInProject(IQuickFix fix, QualifiedSelection selection, Type inspectionType,
            IEnumerable<IInspectionResult> results);

        void FixAll(IQuickFix fix, Type inspectionType, IEnumerable<IInspectionResult> results);

        bool HasQuickFixes(IInspectionResult inspectionResult);
    }
}
