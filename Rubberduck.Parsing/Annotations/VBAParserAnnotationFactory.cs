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
            _creators.Add(AnnotationType.Folder.ToString().ToUpperInvariant(), typeof(FolderAnnotation));
        }

        public IAnnotation Create(VBAParser.AnnotationContext context, QualifiedSelection qualifiedSelection)
        {
            string annotationName = context.annotationName().GetText();
            List<string> parameters = new List<string>();
            var argList = context.annotationArgList();
            if (argList != null)
            {
                parameters.AddRange(argList.annotationArg().Select(arg => arg.GetText()));
            }
            Type annotationCLRType = null;
            if (_creators.TryGetValue(annotationName.ToUpperInvariant(), out annotationCLRType))
            {
                return (IAnnotation)Activator.CreateInstance(annotationCLRType, qualifiedSelection, parameters);
            }
            return null;
        }
    }
}
