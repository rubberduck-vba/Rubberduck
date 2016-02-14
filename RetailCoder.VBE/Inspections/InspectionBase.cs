using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public abstract class InspectionBase : IInspection
    {
        protected readonly RubberduckParserState State;
        private readonly CodeInspectionSeverity _defaultSeverity;

        protected InspectionBase(RubberduckParserState state, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning)
        {
            State = state;
            _defaultSeverity = defaultSeverity;
            Severity = _defaultSeverity;
        }

        /// <summary>
        /// Gets a value the severity level to reset to, the "factory default" setting.
        /// </summary>
        public CodeInspectionSeverity DefaultSeverity { get { return _defaultSeverity; } }

        /// <summary>
        /// Gets a localized string representing a short name/description for the inspection.
        /// </summary>
        public abstract string Description { get; }

        /// <summary>
        /// Gets the type of inspection; used for regrouping inspections.
        /// </summary>
        public abstract CodeInspectionType InspectionType { get; }

        /// <summary>
        /// A method that inspects the parser state and returns all issues it can find.
        /// </summary>
        /// <returns></returns>
        public abstract IEnumerable<InspectionResultBase> GetInspectionResults();

        /// <summary>
        /// The inspection type name, obtained by reflection.
        /// </summary>
        public virtual string Name { get { return GetType().Name; } }

        /// <summary>
        /// Inspection severity level. Can control whether an inspection is enabled.
        /// </summary>
        public CodeInspectionSeverity Severity { get; set; }

        /// <summary>
        /// Meta-information about why an inspection exists.
        /// </summary>
        public virtual string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        
        /// <summary>
        /// Gets a localized string representing the type of inspection.
        /// <see cref="InspectionType"/>
        /// </summary>
        public virtual string InspectionTypeName { get { return InspectionsUI.ResourceManager.GetString(InspectionType.ToString()); } }

        /// <summary>
        /// Gets a string representing the text that must be present in an 
        /// @Ignore annotation to disable the inspection at a given site.
        /// </summary>
        public virtual string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        /// <summary>
        /// Gets all declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> Declarations
        {
            get { return State.AllDeclarations.Where(declaration => !declaration.IsInspectionDisabled(AnnotationName)); }
        }

        /// <summary>
        /// Gets all user declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> UserDeclarations
        {
            get { return State.AllUserDeclarations.Where(declaration => !declaration.IsInspectionDisabled(AnnotationName)); }
        }

        public int CompareTo(IInspection other)
        {
            return string.Compare(InspectionType + Name, other.InspectionType + other.Name, StringComparison.Ordinal);
        }

        public int CompareTo(object obj)
        {
            return CompareTo(obj as IInspection);
        }
    }
}