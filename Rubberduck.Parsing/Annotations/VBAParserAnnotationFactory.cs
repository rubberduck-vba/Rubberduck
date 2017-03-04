using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class VBAParserAnnotationFactory : IAnnotationFactory
    {
        private readonly Dictionary<string, Type> _creators = new Dictionary<string, Type>();

        public VBAParserAnnotationFactory()
        {
            _creators.Add(AnnotationType.TestModule.ToString().ToUpperInvariant(), typeof(TestModuleAnnotation));
            _creators.Add(AnnotationType.ModuleInitialize.ToString().ToUpperInvariant(), typeof(ModuleInitializeAnnotation));
            _creators.Add(AnnotationType.ModuleCleanup.ToString().ToUpperInvariant(), typeof(ModuleCleanupAnnotation));
            _creators.Add(AnnotationType.TestMethod.ToString().ToUpperInvariant(), typeof(TestMethodAnnotation));
            _creators.Add(AnnotationType.TestInitialize.ToString().ToUpperInvariant(), typeof(TestInitializeAnnotation));
            _creators.Add(AnnotationType.TestCleanup.ToString().ToUpperInvariant(), typeof(TestCleanupAnnotation));
            _creators.Add(AnnotationType.Ignore.ToString().ToUpperInvariant(), typeof(IgnoreAnnotation));
            _creators.Add(AnnotationType.IgnoreModule.ToString().ToUpperInvariant(), typeof(IgnoreModuleAnnotation));
            _creators.Add(AnnotationType.IgnoreTest.ToString().ToUpperInvariant(), typeof(IgnoreTestAnnotation));
            _creators.Add(AnnotationType.Folder.ToString().ToUpperInvariant(), typeof(FolderAnnotation));
            _creators.Add(AnnotationType.NoIndent.ToString().ToUpperInvariant(), typeof(NoIndentAnnotation));
            _creators.Add(AnnotationType.Interface.ToString().ToUpperInvariant(), typeof(InterfaceAnnotation));
        }

        public IAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection)
        {
            var annotationName = context.annotationName().GetText();
            var parameters = AnnotationParametersFromContext(context);
            return CreateAnnotation(annotationName, parameters, qualifiedSelection);
        }

            private static List<string> AnnotationParametersFromContext(VBAParser.AnnotationContext context)
            {
                var parameters = new List<string>();
                var argList = context.annotationArgList();
                if (argList != null)
                {
                    parameters.AddRange(argList.annotationArg().Select(arg => arg.GetText()));
                }
                return parameters;
            }

            private IAnnotation CreateAnnotation(string annotationName, List<string> parameters, QualifiedSelection qualifiedSelection)
            {
                Type annotationClrType;
                if (_creators.TryGetValue(annotationName.ToUpperInvariant(), out annotationClrType))
                {
                    return (IAnnotation)Activator.CreateInstance(annotationClrType, qualifiedSelection, parameters);
                }
                return null;
            }
    }
}
