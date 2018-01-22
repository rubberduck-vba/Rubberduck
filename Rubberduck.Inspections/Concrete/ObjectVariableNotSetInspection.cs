using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
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
                    if (!IsIgnoringInspectionResultFor(reference, AnnotationName) 
                        && VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(reference, State))
                    {
                        result.Add(reference);
                    }
                }
            }

            return result;
        }
    }
}
