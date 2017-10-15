﻿using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute after executing each unit test in a test module.
    /// </summary>
    public sealed class TestCleanupAnnotation : AnnotationBase
    {
        public TestCleanupAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestCleanup, qualifiedSelection)
        {
        }
    }
}
