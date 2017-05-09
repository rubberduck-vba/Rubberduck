using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
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
            Name = GetType().Name;
        }

        /// <summary>
        /// Gets a value the severity level to reset to, the "factory default" setting.
        /// </summary>
        public CodeInspectionSeverity DefaultSeverity => _defaultSeverity;

        /// <summary>
        /// Gets a localized string representing a short name/description for the inspection.
        /// </summary>
        public virtual string Description => InspectionsUI.ResourceManager.GetString(Name + "Name", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets the type of inspection; used for regrouping inspections.
        /// </summary>
        public abstract CodeInspectionType InspectionType { get; }

        /// <summary>
        /// A method that inspects the parser state and returns all issues it can find.
        /// </summary>
        /// <returns></returns>
        public abstract IEnumerable<IInspectionResult> GetInspectionResults();

        /// <summary>
        /// The inspection type name, obtained by reflection.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Inspection severity level. Can control whether an inspection is enabled.
        /// </summary>
        public CodeInspectionSeverity Severity { get; set; }

        /// <summary>
        /// Meta-information about why an inspection exists.
        /// </summary>
        public virtual string Meta => InspectionsUI.ResourceManager.GetString(Name + "Meta", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets a localized string representing the type of inspection.
        /// <see cref="InspectionType"/>
        /// </summary>
        public virtual string InspectionTypeName => InspectionsUI.ResourceManager.GetString(InspectionType.ToString(), CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets a string representing the text that must be present in an 
        /// @Ignore annotation to disable the inspection at a given site.
        /// </summary>
        public virtual string AnnotationName => Name.Replace("Inspection", string.Empty);

        /// <summary>
        /// Gets all declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> Declarations
        {
            get { return State.AllDeclarations.Where(declaration => !IsIgnoringInspectionResultFor(declaration, AnnotationName)); }
        }

        /// <summary>
        /// Gets all user declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> UserDeclarations
        {
            get { return State.AllUserDeclarations.Where(declaration => !IsIgnoringInspectionResultFor(declaration, AnnotationName)); }
        }

        protected virtual IEnumerable<Declaration> BuiltInDeclarations
        {
            get { return State.AllDeclarations.Where(declaration => !declaration.IsUserDefined); }
        }

        protected bool IsIgnoringInspectionResultFor(QualifiedModuleName module, int line)
        {
            var annotations = State.GetModuleAnnotations(module).ToList();

            if (State.GetModuleAnnotations(module) == null)
            {
                return false;
            }

            // VBE 1-based indexing
            for (var i = line; i >= 1; i--)
            {
                var annotation = annotations.SingleOrDefault(a => a.QualifiedSelection.Selection.StartLine == i);
                var ignoreAnnotation = annotation as IgnoreAnnotation;
                var ignoreModuleAnnotation = annotation as IgnoreModuleAnnotation;

                if (ignoreAnnotation?.InspectionNames.Contains(AnnotationName) == true)
                {
                    return true;
                }

                if (ignoreModuleAnnotation != null &&
                    (ignoreModuleAnnotation.InspectionNames.Contains(AnnotationName) ||
                     !ignoreModuleAnnotation.InspectionNames.Any()))
                {
                    return true;
                }
            }

            return false;
        }

        protected bool IsIgnoringInspectionResultFor(Declaration declaration, string inspectionName)
        {
            var module = Declaration.GetModuleParent(declaration);
            if (module == null) { return false; }

            var isIgnoredAtModuleLevel = module.Annotations
                    .Any(annotation => annotation.AnnotationType == AnnotationType.IgnoreModule
                                       && ((IgnoreModuleAnnotation) annotation).IsIgnored(inspectionName));


            if (declaration.DeclarationType == DeclarationType.Parameter)
            {
                return isIgnoredAtModuleLevel || declaration.ParentDeclaration.Annotations.Any(annotation =>
                    annotation.AnnotationType == AnnotationType.Ignore
                    && ((IgnoreAnnotation)annotation).IsIgnored(inspectionName));
            }

            return isIgnoredAtModuleLevel || declaration.Annotations.Any(annotation =>
                annotation.AnnotationType == AnnotationType.Ignore
                && ((IgnoreAnnotation)annotation).IsIgnored(inspectionName));
        }

        protected bool IsIgnoringInspectionResultFor(IdentifierReference reference, string inspectionName)
        {
            return reference != null && reference.IsIgnoringInspectionResultFor(inspectionName);
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