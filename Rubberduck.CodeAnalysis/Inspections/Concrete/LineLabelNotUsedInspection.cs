using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies line labels that are never referenced, and therefore superfluous.
    /// </summary>
    /// <why>
    /// Line labels are useful for GoTo, GoSub, Resume, and On Error statements; but the intent of a line label
    /// can be confusing if it isn't referenced by any such instruction.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// '    On Error GoTo ErrHandler ' (commented-out On Error statement leaves line label unreferenced)
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class LineLabelNotUsedInspection : DeclarationInspectionBase
    {
        public LineLabelNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.LineLabel)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration != null
                   && !declaration.IsWithEvents
                   && declaration.Context is VBAParser.IdentifierStatementLabelContext
                   && declaration.References.All(reference => reference.IsAssignment);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.IdentifierNotUsedInspection, 
                declarationType, 
                declarationName);
        }
    }
}
