using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>Locates instances of member calls made against the result of a Range.Find/FindNext/FindPrevious method, without prior validation.</summary>
    /// <reference name="Excel" />
    /// <why>
    /// Range.Find methods return a Range object reference that refers to the cell containing the search string;
    /// this object reference will be Nothing if the search didn't turn up any results, and a member call against Nothing will raise run-time error 91.
    /// </why>
    /// <example hasResults="true">
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
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Sheet1.Range("A:A").Find("Test")
    ///     If Not foo Is Nothing Then
    ///         MsgBox foo.Address ' Range.Address member call is safe.
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    [RequiredLibrary("Excel")]
    public class ExcelMemberMayReturnNothingInspection : MemberAccessMayReturnNothingInspectionBase
    {
        public ExcelMemberMayReturnNothingInspection(RubberduckParserState state) : base(state) { }

        private static readonly List<string> ExcelMembers = new List<string>
        {
            "Range.Find",
            "Range.FindNext",
            "Range.FindPrevious"
        };

        public override List<Declaration> MembersUnderTest => BuiltInDeclarations
            .Where(decl => decl.ProjectName.Equals("Excel") && ExcelMembers.Any(member => decl.QualifiedName.ToString().EndsWith(member)))
            .ToList();

        public override string ResultTemplate => InspectionResults.ExcelMemberMayReturnNothingInspection;
    }
}
