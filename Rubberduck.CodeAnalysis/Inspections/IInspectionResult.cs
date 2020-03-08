using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections
{
    public interface IInspectionResult : IComparable<IInspectionResult>, IComparable
    {
        string Description { get; }
        QualifiedSelection QualifiedSelection { get; }
        QualifiedMemberName? QualifiedMemberName { get; }
        IInspection Inspection { get; }
        Declaration Target { get; }
        ParserRuleContext Context { get; }
        ICollection<string> DisabledQuickFixes { get; }
        bool ChangesInvalidateResult(ICollection<QualifiedModuleName> modifiedModules);
    }

    public interface IWithInspectionResultProperties<T>
    {
        T Properties { get; }
    }
}
