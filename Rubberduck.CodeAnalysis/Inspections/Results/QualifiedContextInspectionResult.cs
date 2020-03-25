using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Results
{
    internal class QualifiedContextInspectionResult : InspectionResultBase
    {
        public QualifiedContextInspectionResult(
            IInspection inspection, 
            string description, 
            QualifiedContext context,
            ICollection<string> disabledQuickFixes = null) :
            base(inspection,
                 description,
                 context.ModuleName,
                 context.Context,
                 null,
                 new QualifiedSelection(context.ModuleName, context.Context.GetSelection()),
                 context.MemberName,
                 disabledQuickFixes)
        {}
    }

    internal class QualifiedContextInspectionResult<T> : QualifiedContextInspectionResult, IWithInspectionResultProperties<T>
    {
        public QualifiedContextInspectionResult(
            IInspection inspection, 
            string description, 
            QualifiedContext context,
            T properties,
            ICollection<string> disabledQuickFixes = null) :
            base(
                inspection,
                description,
                context,
                disabledQuickFixes)
        {
            Properties = properties;
        }

        public T Properties { get; }
    }
}
