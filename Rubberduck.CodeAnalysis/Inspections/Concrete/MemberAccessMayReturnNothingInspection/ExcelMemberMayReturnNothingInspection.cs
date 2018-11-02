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
