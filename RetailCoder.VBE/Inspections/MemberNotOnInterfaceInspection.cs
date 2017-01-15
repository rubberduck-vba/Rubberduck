using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class MemberNotOnInterfaceInspection : InspectionBase
    {
        private static readonly List<Type> InterestingTypes = new List<Type>
        {
            typeof(VBAParser.MemberAccessExprContext),
            typeof(VBAParser.WithMemberAccessExprContext),
            typeof(VBAParser.DictionaryAccessExprContext),
            typeof(VBAParser.WithDictionaryAccessExprContext)
        }; 

        public MemberNotOnInterfaceInspection(RubberduckParserState state, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning)
            : base(state, defaultSeverity)
        {
        }

        public override string Meta { get { return InspectionsUI.MemberNotOnInterfaceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MemberNotOnInterfaceInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var targets = Declarations.Where(decl => decl.AsTypeDeclaration != null && 
                                                     decl.AsTypeDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                     ((ClassModuleDeclaration)decl.AsTypeDeclaration).IsExtensible &&
                                                     decl.References.Any(usage => InterestingTypes.Contains(usage.Context.Parent.GetType())))
                                            .ToList();

            //Unfortunately finding implemented members is fairly expensive, so group by the return type.
            return (from access in targets.GroupBy(t => t.AsTypeDeclaration)
                let typeDeclaration = access.Key
                let typeMembers = new HashSet<string>(BuiltInDeclarations.Where(d => d.ParentDeclaration != null && d.ParentDeclaration.Equals(typeDeclaration))
                                                                         .Select(d => d.IdentifierName)
                                                                         .Distinct())
                from references in access.Select(usage => usage.References.Where(r => InterestingTypes.Contains(r.Context.Parent.GetType())))
                from reference in references.Where(r => !r.IsInspectionDisabled(AnnotationName))
                let identifier = reference.Context.Parent.GetChild(reference.Context.Parent.ChildCount - 1)
                where !typeMembers.Contains(identifier.GetText())
                let pseudoDeclaration = CreatePseudoDeclaration((ParserRuleContext) identifier, reference)
                where !pseudoDeclaration.Annotations.Any()
                select new MemberNotOnInterfaceInspectionResult(this, pseudoDeclaration, (ParserRuleContext) identifier, typeDeclaration))
                                                               .Cast<InspectionResultBase>().ToList();
        }

        //Builds a throw-away Declaration for the indentifiers found by the inspection. These aren't (and shouldn't be) created by the parser.
        //Used to pass to the InspectionResult to make it selectable.
        private static Declaration CreatePseudoDeclaration(ParserRuleContext context, IdentifierReference reference)
        {
            return new Declaration(
                new QualifiedMemberName(reference.QualifiedModuleName, context.GetText()),
                null, 
                null, 
                string.Empty, 
                string.Empty, 
                false, 
                false,
                Accessibility.Implicit, 
                DeclarationType.Variable, 
                context,
                context.GetSelection(), 
                false, 
                null, 
                true,
                null,
                null, 
                true);
        }
    }
}
