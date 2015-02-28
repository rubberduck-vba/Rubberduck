using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Inspections
{
    public interface IInspector
    {
        IList<ICodeInspectionResult> FindIssues(VBProject project);
    }
}
