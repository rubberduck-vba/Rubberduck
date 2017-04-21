using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SynchronizeAttributesQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MissingAnnotationInspection),
            typeof(MissingAttributeInspection),
        };

        public SynchronizeAttributesQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public void Fix(IInspectionResult result)
        {
            var context = result.Context;
            var memberName = result.QualifiedMemberName;

            var attributeContext = context as VBAParser.AttributeStmtContext;
            if (attributeContext != null)
            {
                Fix(memberName, attributeContext);
                return;
            }

            var annotationContext = context as VBAParser.AnnotationContext;
            if (annotationContext != null)
            {
                Fix(memberName, annotationContext);
                return;
            }
        }

        /// <summary>
        /// Adds an annotation to match given attribute.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName? memberName, VBAParser.AttributeStmtContext context)
        {
            
        }

        /// <summary>
        /// Adds an attribute to match given annotation.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName? memberName, VBAParser.AnnotationContext context)
        {
            
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SynchronizeAttributesQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}