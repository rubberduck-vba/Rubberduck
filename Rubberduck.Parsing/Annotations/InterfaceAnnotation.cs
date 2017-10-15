﻿using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used to mark a class module as an interface, so that Rubberduck treats it as such even if it's not implemented in any opened project.
    /// </summary>
    public sealed class InterfaceAnnotation : AnnotationBase
    {
        public InterfaceAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Interface, qualifiedSelection)
        {
        }
    }
}