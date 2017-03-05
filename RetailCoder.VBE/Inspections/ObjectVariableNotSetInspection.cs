using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.ObjectVariableNotSetInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObjectVariableNotSetInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var interestingDeclarations =
                State.AllUserDeclarations.Where(item =>
                        !item.IsSelfAssigned &&
                        !item.IsArray &&
                        !SymbolList.ValueTypes.Contains(item.AsTypeName) &&
                        (item.AsTypeDeclaration == null || (!ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration) &&
                        item.AsTypeDeclaration.DeclarationType != DeclarationType.Enumeration &&
                        item.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)) &&
                        (item.DeclarationType == DeclarationType.Variable ||
                         item.DeclarationType == DeclarationType.Parameter));

            var interestingMembers =
                State.AllUserDeclarations.Where(item =>
                    (item.DeclarationType == DeclarationType.Function || item.DeclarationType == DeclarationType.PropertyGet)
                    && !item.IsArray
                    && item.IsTypeSpecified
                    && !SymbolList.ValueTypes.Contains(item.AsTypeName) 
                    && (item.AsTypeDeclaration == null // null if unresolved (e.g. in unit tests)
                        || (item.AsTypeDeclaration.DeclarationType != DeclarationType.Enumeration && item.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType 
                            && item.AsTypeDeclaration != null 
                            && !ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration))));

            var interestingReferences = interestingDeclarations
                    .Union(interestingMembers.SelectMany(item =>
                        item.References.Where(reference => reference.ParentScoping.Equals(item) && reference.IsAssignment)
                    .Select(reference => reference.Declaration)))
                    .SelectMany(declaration =>
                        declaration.References.Where(reference =>
                        {
                            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
                            return reference.IsAssignment && letStmtContext != null && letStmtContext.LET() == null;
                        })
                    );


            return interestingReferences.Select(reference => new ObjectVariableNotSetInspectionResult(this, reference));
        }
    }
}
