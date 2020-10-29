using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using NLog;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class InspectionBase : IInspection
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        protected readonly ILogger Logger;

        protected InspectionBase(IDeclarationFinderProvider declarationFinderProvider)
        {
            Logger = LogManager.GetLogger(GetType().FullName);

            _declarationFinderProvider = declarationFinderProvider;
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
        public virtual string InspectionTypeName => Resources.Inspections.InspectionsUI.ResourceManager.GetString($"CodeInspectionSettings_{InspectionType}", CultureInfo.CurrentUICulture);

        /// <summary>
        /// Gets a string representing the text that must be present in an 
        /// @Ignore annotation to disable the inspection at a given site.
        /// </summary>
        public virtual string AnnotationName => Name.Replace("Inspection", string.Empty);

        public int CompareTo(IInspection other) => string.Compare(InspectionType + Name, other.InspectionType + other.Name, StringComparison.Ordinal);
        public int CompareTo(object obj) => CompareTo(obj as IInspection);

        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder);
        protected abstract IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder);

        /// <summary>
        /// A method that inspects the parser state and returns all issues it can find.
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        public IEnumerable<IInspectionResult> GetInspectionResults(CancellationToken token)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            var finder = _declarationFinderProvider.DeclarationFinder;
            var result = DoGetInspectionResults(finder)
                .Where(ir => !ir.IsIgnoringInspectionResult(finder))
                .ToList();
            stopwatch.Stop();
            Logger.Trace("Intercepted invocation of '{0}.{1}' returned {2} objects.", GetType().Name, nameof(DoGetInspectionResults), result.Count);
            Logger.Trace("Intercepted invocation of '{0}.{1}' ran for {2}ms", GetType().Name, nameof(DoGetInspectionResults), stopwatch.ElapsedMilliseconds);
            return result;
        }

        /// <summary>
        /// A method that inspects the parser state and returns all issues in it can find in a module.
        /// </summary>
        /// <param name="module">The module for which to get inspection results</param>
        /// <param name="token"></param>
        /// <returns></returns>
        public IEnumerable<IInspectionResult> GetInspectionResults(QualifiedModuleName module, CancellationToken token)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            var finder = _declarationFinderProvider.DeclarationFinder;
            var result = DoGetInspectionResults(module, finder)
                .Where(ir => !ir.IsIgnoringInspectionResult(finder))
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