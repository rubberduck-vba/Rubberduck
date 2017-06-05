using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
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

        private readonly IDictionary<string, string> _attributeNames;

        public SynchronizeAttributesQuickFix(RubberduckParserState state)
        {
            _state = state;
            _attributeNames = typeof(AnnotationType).GetFields()
                .Where(field => field.GetCustomAttributes(typeof (AttributeAnnotationAttribute), true).Any())
                .Select(a => new { AnnotationName = a.Name, a.GetCustomAttributes(typeof (AttributeAnnotationAttribute), true).Cast<AttributeAnnotationAttribute>().FirstOrDefault()?.AttributeName})
                .ToDictionary(a => a.AnnotationName, a => a.AttributeName);
        }

        public void Fix(IInspectionResult result)
        {
            var context = result.Context;
            if (result.QualifiedMemberName != null)
            {
                var memberName = result.QualifiedMemberName.Value;

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
            else
            {
                var moduleName = result.QualifiedSelection.QualifiedName;

                var attributeContext = context as VBAParser.AttributeStmtContext;
                if(attributeContext != null)
                {
                    Fix(moduleName, attributeContext);
                    return;
                }

                var annotationContext = context as VBAParser.AnnotationContext;
                if(annotationContext != null)
                {
                    Fix(moduleName, annotationContext);
                    return;
                }
            }
        }

        /// <summary>
        /// Adds an annotation to match given attribute.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName memberName, VBAParser.AttributeStmtContext context)
        {

        }

        private void Fix(QualifiedModuleName moduleName, VBAParser.AttributeStmtContext context)
        {
            
        }

        private void Fix(QualifiedModuleName moduleName, VBAParser.AnnotationContext context)
        {
            var annotationName = Identifier.GetName(context.annotationName().unrestrictedIdentifier());
            var annotationType = context.AnnotationType;
            var attributeName = _attributeNames[annotationName];

            var attributeInstruction = GetAttributeInstruction(context, attributeName, annotationType);

            var rewriter = _state.GetAttributeRewriter(moduleName);
        }

        /// <summary>
        /// Adds an attribute to match given annotation.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName memberName, VBAParser.AnnotationContext context)
        {
            Debug.Assert(context.AnnotationType.HasFlag(AnnotationType.MemberAnnotation));

            var annotationName = Identifier.GetName(context.annotationName().unrestrictedIdentifier());
            var annotationType = context.AnnotationType;
            var attributeName =  memberName.MemberName + "." + _attributeNames[annotationName];

            var attributeInstruction = GetAttributeInstruction(context, attributeName, annotationType);
            var insertPosition = FindInsertPosition(context);

            var rewriter = _state.GetAttributeRewriter(memberName.QualifiedModuleName);
            rewriter.InsertBefore(insertPosition, attributeInstruction);
        }

        private int FindInsertPosition(VBAParser.AnnotationContext context)
        {
            return (context.AnnotatedContext as IAnnotatedContext)?.AttributeTokenIndex ?? 1;
        }

        private string GetAttributeInstruction(VBAParser.AnnotationContext context, string attributeName, AnnotationType annotationType)
        {
            string attributeInstruction = string.Empty;

            if (annotationType.HasFlag(AnnotationType.ModuleAnnotation))
            {
                switch (annotationType)
                {
                    case AnnotationType.Exposed:
                        attributeInstruction = $"Attribute {attributeName} = True\n";
                        break;
                    case AnnotationType.PredeclaredId:
                        attributeInstruction = $"Attribute {attributeName} = True\n";
                        break;
                }
            }
            else if (annotationType.HasFlag(AnnotationType.MemberAnnotation))
            {
                switch (annotationType)
                {
                    case AnnotationType.Description:
                        var description = context.annotationArgList().annotationArg().FirstOrDefault()?.GetText() ?? string.Empty;
                        description = description.StartsWith("\"") && description.EndsWith("\"")
                            ? description
                            : $"\"{description}\"";

                        attributeInstruction = $"Attribute {attributeName} = \"{description}\"\n";
                        break;
                    case AnnotationType.DefaultMember:
                        attributeInstruction = $"Attribute {attributeName} = 0";
                        break;
                    case AnnotationType.Enumerator:
                        attributeInstruction = $"Attribute {attributeName} = -4";
                        break;
                }
            }

            return attributeInstruction;
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