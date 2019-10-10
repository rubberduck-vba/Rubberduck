﻿using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used to indicate the test engine that a unit test is to be ignored.
    /// </summary>
    public sealed class IgnoreTestAnnotation : AnnotationBase
    {
        public IgnoreTestAnnotation()
            : base("IgnoreTest", AnnotationTarget.Member)
        { }
    }
}