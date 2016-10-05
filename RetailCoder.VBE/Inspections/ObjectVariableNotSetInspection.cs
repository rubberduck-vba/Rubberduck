using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObjectVariableNotSetInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObjectVariableNotSetInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new UseSetKeywordForObjectAssignmentQuickFix(_reference),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, _reference.Declaration.IdentifierName); }
        }
    }

    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.ObjectVariableNotSetInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObjectVariableNotSetInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly IReadOnlyList<string> ValueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.String,
            Tokens.Variant
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var interestingDeclarations =
                State.AllUserDeclarations.Where(item =>
                        !item.IsSelfAssigned &&
                        !item.IsArray &&
                        !ValueTypes.Contains(item.AsTypeName) &&
                        (item.AsTypeDeclaration == null ||
                        item.AsTypeDeclaration.DeclarationType != DeclarationType.Enumeration &&
                        item.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType) &&
                        (item.DeclarationType == DeclarationType.Variable ||
                         item.DeclarationType == DeclarationType.Parameter));

            var interestingMembers =
                State.AllUserDeclarations.Where(item =>
                    (item.DeclarationType == DeclarationType.Function || item.DeclarationType == DeclarationType.PropertyGet)
                    && !item.IsArray
                    && item.IsTypeSpecified
                    && !ValueTypes.Contains(item.AsTypeName));

            var interestingReferences = interestingDeclarations
                    .Union(interestingMembers.SelectMany(item =>
                        item.References.Where(reference =>
                            reference.ParentScoping == item && reference.IsAssignment
                        ).Select(reference => reference.Declaration))
                    )
                    .SelectMany(declaration =>
                        declaration.References.Where(reference =>
                        {
                            var setStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
                            return reference.IsAssignment && setStmtContext != null && setStmtContext.LET() == null;
                        })
                    );


            return interestingReferences.Select(reference => new ObjectVariableNotSetInspectionResult(this, reference));
        }
    }
}
