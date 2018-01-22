using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) {  }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {

            return InterestingReferences().Select(reference =>
                new IdentifierReferenceInspectionResult(this,
                    string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, reference.Declaration.IdentifierName),
                    State, reference));
        }

        private IEnumerable<IdentifierReference> InterestingReferences()
        {
            var result = new List<IdentifierReference>();
            foreach (var qmn in State.DeclarationFinder.AllModules.Where(m => m.ComponentType != ComponentType.Undefined && m.ComponentType != ComponentType.ComComponent))
            {
                var module = State.DeclarationFinder.ModuleDeclaration(qmn);
                if (module == null || !module.IsUserDefined || IsIgnoringInspectionResultFor(module, AnnotationName))
                {
                    // module isn't user code, or this inspection is ignored at module-level
                    continue;
                }

                foreach (var reference in State.DeclarationFinder.IdentifierReferences(qmn))
                {
                    if (IsIgnoringInspectionResultFor(reference, AnnotationName))
                    {
                        // inspection is ignored at reference level
                        continue;
                    }

                    if (!reference.IsAssignment)
                    {
                        // reference isn't assigning its declaration; not interesting
                        continue;
                    }

                    var setStmtContext = ParserRuleContextHelper.GetParent<VBAParser.SetStmtContext>(reference.Context);
                    if (setStmtContext != null)
                    {
                        // assignment already has a Set keyword
                        // (but is it misplaced? ...hmmm... beyond the scope of *this* inspection though)
                        continue;
                    }

                    var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
                    if (letStmtContext == null)
                    {
                        // we're probably in a For Each loop
                        continue;
                    }

                    var declaration = reference.Declaration;
                    if (declaration.IsArray)
                    {
                        // arrays don't need a Set statement... todo figure out if array items are objects
                        continue;
                    }

                    var isObjectVariable =
                        declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? false;
                    var isVariant = declaration.IsUndeclared || declaration.AsTypeName == Tokens.Variant;
                    if (!isObjectVariable && !isVariant)
                    {
                        continue;
                    }

                    if (isObjectVariable)
                    {
                        var members = State.DeclarationFinder.Members(declaration.AsTypeDeclaration).ToHashSet();
                        var parameters = members.Where(m => m.DeclarationType == DeclarationType.Parameter).Cast<ParameterDeclaration>().ToHashSet();
                        if (members.Any(member => !parameters.Any() || parameters.All(p => p.IsOptional && p.ParentScopeDeclaration.Equals(member)) && member.Attributes.HasDefaultMemberAttribute()))
                        {
                            // assigned declaration has a default parameterless member, which is legally being assigned here.
                            // might be a good idea to flag that default member assignment though...
                            continue;
                        }

                        // assign declaration is an object without a default parameterless (or with all parameters optional) member - LHS needs a 'Set' keyword.
                        result.Add(reference);
                        continue;
                    }
                    
                    // assigned declaration is a variant. we need to know about the RHS of the assignment.

                    var expression = letStmtContext.expression();
                    if (expression == null)
                    {
                        Debug.Assert(false, "RHS expression is empty? What's going on here?");
                    }

                    if (expression is VBAParser.NewExprContext)
                    {
                        // RHS expression is newing up an object reference - LHS needs a 'Set' keyword:
                        result.Add(reference);
                        continue;
                    }

                    var literalExpression = expression as VBAParser.LiteralExprContext;
                    if (literalExpression?.literalExpression()?.literalIdentifier()?.objectLiteralIdentifier() != null)
                    {
                        // RHS is a 'Nothing' token - LHS needs a 'Set' keyword:
                        result.Add(reference);
                        continue;
                    }

                    // todo resolve expression return type

                    var memberRefs = State.DeclarationFinder.IdentifierReferences(reference.ParentScoping.QualifiedName);
                    var lastRef = memberRefs.LastOrDefault(r => !Equals(r, reference) && ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(r.Context) == letStmtContext);
                    if (lastRef?.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? false)
                    {
                        // the last reference in the expression is referring to an object type
                        result.Add(reference);
                        continue;
                    }
                    if (lastRef?.Declaration.AsTypeName == Tokens.Object)
                    {
                        result.Add(reference);
                        continue;
                    }

                    var accessibleDeclarations  = State.DeclarationFinder.GetAccessibleDeclarations(reference.ParentScoping);
                    foreach (var accessibleDeclaration in accessibleDeclarations.Where(d => d.IdentifierName == expression.GetText()))
                    {
                        if (accessibleDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule) || accessibleDeclaration.AsTypeName == Tokens.Object)
                        {
                            result.Add(reference);
                            break;
                        }
                    }
                }
            }

            return result;
        }
    }
}
