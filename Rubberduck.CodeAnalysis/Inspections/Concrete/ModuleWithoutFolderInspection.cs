using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a user module is missing a @Folder Rubberduck annotation.
    /// </summary>
    /// <why>
    /// Modules without a custom @Folder annotation will be grouped under the default folder in the Code Explorer toolwindow.
    /// By specifying a custom @Folder annotation, modules can be organized by functionality rather than simply listed.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Folder("Foo")
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class ModuleWithoutFolderInspection : InspectionBase
    {
        public ModuleWithoutFolderInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var modulesWithoutFolderAnnotation = State.DeclarationFinder.UserDeclarations(Parsing.Symbols.DeclarationType.Module)
                .Where(w => !w.Annotations.Any(pta => pta.Annotation is FolderAnnotation))
                .ToList();

            return modulesWithoutFolderAnnotation
                .Select(declaration =>
                new DeclarationInspectionResult(this, string.Format(InspectionResults.ModuleWithoutFolderInspection, declaration.IdentifierName), declaration));
        }
    }
}
