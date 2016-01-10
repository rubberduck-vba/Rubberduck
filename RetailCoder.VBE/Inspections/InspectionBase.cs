using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public abstract class InspectionBase : IInspection
    {
        protected readonly RubberduckParserState State;
        protected InspectionBase(RubberduckParserState state)
        {
            State = state;
        }

        public abstract string Description { get; }
        public abstract CodeInspectionType InspectionType { get; }
        public abstract IEnumerable<CodeInspectionResultBase> GetInspectionResults();

        public virtual string Name { get { return GetType().Name; } }
        public virtual CodeInspectionSeverity Severity { get; set; }
        public virtual string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public virtual string InspectionTypeName { get { return InspectionsUI.ResourceManager.GetString(InspectionType.ToString()); } }
        public virtual string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        protected virtual IEnumerable<Declaration> Declarations
        {
            get { return State.AllDeclarations.Where(declaration => !declaration.IsInspectionDisabled(AnnotationName)); }
        }

        protected virtual IEnumerable<Declaration> UserDeclarations
        {
            get { return State.AllUserDeclarations.Where(declaration => !declaration.IsInspectionDisabled(AnnotationName)); }
        }
    }
}