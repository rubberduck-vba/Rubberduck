using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class ObjectVariableNotSetInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObjectVariableNotSetInspectionResult(IInspection inspection, IdentifierReference reference)
            :base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new SetObjectVariableQuickFix(_reference), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, _reference.Declaration.IdentifierName); }
        }
    }

    public sealed class SetObjectVariableQuickFix : CodeInspectionQuickFix
    {
        public SetObjectVariableQuickFix(IdentifierReference reference)
            : base(context: reference.Context.Parent.Parent as ParserRuleContext, // ImplicitCallStmt_InStmtContext 
                   selection: new QualifiedSelection(reference.QualifiedModuleName, reference.Selection), 
                   description: InspectionsUI.SetObjectVariableQuickFix)
        {
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        public override void Fix()
        {
            var codeModule = Selection.QualifiedName.Component.CodeModule;
            var codeLine = codeModule.get_Lines(Selection.Selection.StartLine, 1);

            var letStatementLeftSide = Context.GetText();
            var setStatementLeftSide = Tokens.Set + ' ' + letStatementLeftSide;

            var correctLine = codeLine.Replace(letStatementLeftSide, setStatementLeftSide);
            codeModule.ReplaceLine(Selection.Selection.StartLine, correctLine);
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
            return State.AllUserDeclarations
                .Where(item => !ValueTypes.Contains(item.AsTypeName)
                    && !item.IsSelfAssigned
                               && (item.DeclarationType == DeclarationType.Variable
                                   || item.DeclarationType == DeclarationType.Parameter))
                .SelectMany(declaration =>
                    declaration.References.Where(reference =>
                    {
                        var setStmtContext = reference.Context.Parent.Parent.Parent as VBAParser.LetStmtContext;
                        return setStmtContext != null && setStmtContext.LET() == null;
                    }))
                .Select(reference => new ObjectVariableNotSetInspectionResult(this, reference));
        }
    }
}