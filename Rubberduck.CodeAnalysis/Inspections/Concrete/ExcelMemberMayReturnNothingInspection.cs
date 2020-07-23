using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>Locates instances of member calls made against the result of a Range.Find/FindNext/FindPrevious method, without prior validation.</summary>
    /// <reference name="Excel" />
    /// <why>
    /// Range.Find methods return a Range object reference that refers to the cell containing the search string;
    /// this object reference will be Nothing if the search didn't turn up any results, and a member call against Nothing will raise run-time error 91.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Sheet1.Range("A:A").Find("Test") ' foo is Nothing if there are no results
    ///     MsgBox foo.Address ' Range.Address member call should be flagged.
    /// 
    ///     Dim rowIndex As Range
    ///     rowIndex = Sheet1.Range("A:A").Find("Test").Row ' Range.Row member call should be flagged.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Sheet1.Range("A:A").Find("Test")
    ///     If Not foo Is Nothing Then
    ///         MsgBox foo.Address ' Range.Address member call is safe.
    ///     End If
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal class ExcelMemberMayReturnNothingInspection : MemberAccessMayReturnNothingInspectionBase
    {
        public ExcelMemberMayReturnNothingInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        private static readonly List<(string className, string memberName)> ExcelMembers = new List<(string className, string memberName)>
        {
            ("Range","Find"),
            ("Range","FindNext"),
            ("Range","FindPrevious")
        };

        public override IEnumerable<Declaration> MembersUnderTest(DeclarationFinder finder)
        {
            var excel = finder.Projects
                .SingleOrDefault(item => !item.IsUserDefined
                                         && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var memberModules = ExcelMembers
                .Select(tpl => tpl.className)
                .Distinct()
                .Select(className => finder.FindClassModule(className, excel, true))
                .OfType<ModuleDeclaration>();

            return memberModules
                .SelectMany(module => module.Members)
                .Where(member => ExcelMembers.Contains((member.ComponentName, member.IdentifierName)));
        }

        public override string ResultTemplate => InspectionResults.ExcelMemberMayReturnNothingInspection;
    }
}
