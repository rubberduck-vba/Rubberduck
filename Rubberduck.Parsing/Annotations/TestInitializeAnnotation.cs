﻿using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute before executing each unit test in a test module.
    /// </summary>
    public sealed class TestInitializeAnnotation : AnnotationBase
    {
        public TestInitializeAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestInitialize, qualifiedSelection)
        {
        }
    }
}
