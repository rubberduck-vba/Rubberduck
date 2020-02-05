using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class QualifiedContextInspectionResult : InspectionResultBase
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

    public class QualifiedContextInspectionResult<TProperties> : QualifiedContextInspectionResult
    {
        private readonly TProperties _properties;

        public QualifiedContextInspectionResult(
            IInspection inspection, 
            string description, 
            QualifiedContext context,
            TProperties properties,
            ICollection<string> disabledQuickFixes = null) :
            base(
                inspection,
                description,
                context,
                disabledQuickFixes)
        {
            _properties = properties;
        }

        public override T Properties<T>()
        {
            if (_properties is T properties)
            {
                return properties;
            }

            return base.Properties<T>();
        }
    }
}
