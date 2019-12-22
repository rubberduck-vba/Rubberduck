using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Global' declaration statements.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete; use the 'Public' keyword instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Global Foo As Long
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Foo As Long
    /// ]]>
    /// </example>
    public sealed class ObsoleteGlobalInspection : InspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return from item in UserDeclarations
                   where item.Accessibility == Accessibility.Global && item.Context != null
                   select new DeclarationInspectionResult(this,
                       string.Format(InspectionResults.ObsoleteGlobalInspection, item.DeclarationType.ToLocalizedString(), item.IdentifierName), item);
        }
    }
}
