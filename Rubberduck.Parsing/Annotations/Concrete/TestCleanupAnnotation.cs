﻿using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Marks a method that the test engine will execute after executing each unit test in a test module.
    /// </summary>
    public sealed class TestCleanupAnnotation : AnnotationBase
    {
        public TestCleanupAnnotation()
            : base("TestCleanup", AnnotationTarget.Member)
        {
        }
    }
}
