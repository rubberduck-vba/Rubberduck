using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IInspectionResult : IComparable<IInspectionResult>, IComparable
    {
        string Description { get; }
        QualifiedSelection QualifiedSelection { get; }
        QualifiedMemberName? QualifiedMemberName { get; }
        IInspection Inspection { get; }
        Declaration Target { get; }
        ParserRuleContext Context { get; }
        IDictionary<string, string> Properties { get; }
    }
}
