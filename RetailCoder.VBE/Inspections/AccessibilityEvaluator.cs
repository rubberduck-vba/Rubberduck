using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    public static class AccessibilityEvaluator
    {
        public static IEnumerable<Declaration> GetDeclarationsAccessibleToScope(Declaration target, IEnumerable<Declaration> declarations)
        {
            if (target == null) { return Enumerable.Empty<Declaration>(); }

            return declarations
                .Where(candidateDeclaration =>
                (
                       IsDeclarationInTheSameProcedure(candidateDeclaration, target)
                    || IsDeclarationChildOfTheScope(candidateDeclaration, target)
                    || IsModuleLevelDeclarationOfTheScope(candidateDeclaration, target)
                    || IsProjectGlobalDeclaration(candidateDeclaration, target)
                 )).Distinct();
        }

        private static bool IsDeclarationInTheSameProcedure(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ParentScope == scopingDeclaration.ParentScope;
        }

        private static bool IsDeclarationChildOfTheScope(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return scopingDeclaration == candidateDeclaration.ParentDeclaration;
        }

        private static bool IsModuleLevelDeclarationOfTheScope(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            if (candidateDeclaration.ParentDeclaration == null)
            {
                return false;
            }
            return candidateDeclaration.ComponentName == scopingDeclaration.ComponentName
                    && !IsDeclaredWithinMethodOrProperty(candidateDeclaration.ParentDeclaration.Context);
        }

        private static bool IsProjectGlobalDeclaration(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ProjectName == scopingDeclaration.ProjectName
                && !(candidateDeclaration.ParentScopeDeclaration is ClassModuleDeclaration)
                && (candidateDeclaration.Accessibility == Accessibility.Public
                    || ((candidateDeclaration.Accessibility == Accessibility.Implicit)
                        && (candidateDeclaration.ParentScopeDeclaration is ProceduralModuleDeclaration)));
        }

        private static bool IsDeclaredWithinMethodOrProperty(RuleContext procedureContextCandidate)
        {
            if (procedureContextCandidate == null) { return false; }

            return (procedureContextCandidate is VBAParser.SubStmtContext)
                || (procedureContextCandidate is VBAParser.FunctionStmtContext)
                || (procedureContextCandidate is VBAParser.PropertyLetStmtContext)
                || (procedureContextCandidate is VBAParser.PropertyGetStmtContext)
                || (procedureContextCandidate is VBAParser.PropertySetStmtContext);
        }
    }
}
