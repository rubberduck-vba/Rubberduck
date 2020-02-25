using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies line labels that are never referenced, and therefore superfluous.
    /// </summary>
    /// <why>
    /// Line labels are useful for GoTo, GoSub, Resume, and On Error statements; but the intent of a line label
    /// can be confusing if it isn't referenced by any such instruction.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// '    On Error GoTo ErrHandler ' (commented-out On Error statement leaves line label unreferenced)
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class LineLabelNotUsedInspection : DeclarationInspectionBase
    {
        public LineLabelNotUsedInspection(RubberduckParserState state) 
            : base(state, DeclarationType.LineLabel)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration != null
                   && !declaration.IsWithEvents
                   && declaration.Context is IdentifierStatementLabelContext
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
