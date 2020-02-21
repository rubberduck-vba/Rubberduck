using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Diagnostics;
using System.Threading;
using NLog;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class InspectionBase : IInspection
    {
        protected readonly RubberduckParserState State;
        protected readonly IDeclarationFinderProvider DeclarationFinderProvider;

        protected readonly ILogger Logger;

        protected InspectionBase(RubberduckParserState state)
        {
            Logger = LogManager.GetLogger(GetType().FullName);

            State = state;
            DeclarationFinderProvider = state;
            Name = GetType().Name;
        }

        /// <summary>
        /// Gets a localized string representing a short name/description for the inspection.
        /// </summary>
        public virtual string Description => Resources.Inspections.InspectionNames.ResourceManager.GetString($"{Name}", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets the type of inspection; used for regrouping inspections.
        /// </summary>
        public CodeInspectionType InspectionType { get; set; } = CodeInspectionType.CodeQualityIssues;

        /// <summary>
        /// The inspection type name, obtained by reflection.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Inspection severity level. Can control whether an inspection is enabled.
        /// </summary>
        public CodeInspectionSeverity Severity { get; set; } = CodeInspectionSeverity.Warning;

        /// <summary>
        /// Meta-information about why an inspection exists.
        /// </summary>
        public virtual string Meta => Resources.Inspections.InspectionInfo.ResourceManager.GetString($"{Name}", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets a localized string representing the type of inspection.
        /// <see cref="InspectionType"/>
        /// </summary>
        public virtual string InspectionTypeName => Resources.Inspections.InspectionsUI.ResourceManager.GetString($"CodeInspectionSettings_{InspectionType.ToString()}", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets a string representing the text that must be present in an 
        /// @Ignore annotation to disable the inspection at a given site.
        /// </summary>
        public virtual string AnnotationName => Name.Replace("Inspection", string.Empty);

        /// <summary>
        /// Gets all declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> Declarations => DeclarationFinderProvider
            .DeclarationFinder
            .AllDeclarations
            .Where(declaration => !declaration.IsIgnoringInspectionResultFor(AnnotationName));

        /// <summary>
        /// Gets all user declarations in the parser state without an @Ignore annotation for this inspection.
        /// </summary>
        protected virtual IEnumerable<Declaration> UserDeclarations => DeclarationFinderProvider
            .DeclarationFinder
            .AllUserDeclarations
            .Where(declaration => !declaration.IsIgnoringInspectionResultFor(AnnotationName));

        protected virtual IEnumerable<Declaration> BuiltInDeclarations => DeclarationFinderProvider
            .DeclarationFinder
            .AllBuiltInDeclarations;

        public int CompareTo(IInspection other) => string.Compare(InspectionType + Name, other.InspectionType + other.Name, StringComparison.Ordinal);
        public int CompareTo(object obj) => CompareTo(obj as IInspection);

        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults();

        /// <summary>
        /// A method that inspects the parser state and returns all issues it can find.
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        public IEnumerable<IInspectionResult> GetInspectionResults(CancellationToken token)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            var declarationFinder = DeclarationFinderProvider.DeclarationFinder;
            var result = DoGetInspectionResults()
                .Where(ir => !ir.IsIgnoringInspectionResult(declarationFinder))
                .ToList();
            stopwatch.Stop();
            Logger.Trace("Intercepted invocation of '{0}.{1}' returned {2} objects.", GetType().Name, nameof(DoGetInspectionResults), result.Count);
            Logger.Trace("Intercepted invocation of '{0}.{1}' ran for {2}ms", GetType().Name, nameof(DoGetInspectionResults), stopwatch.ElapsedMilliseconds);
            return result;
        }

        public virtual bool ChangesInvalidateResult(IInspectionResult result, ICollection<QualifiedModuleName> modifiedModules)
        {
            return true;
        }
    }
}