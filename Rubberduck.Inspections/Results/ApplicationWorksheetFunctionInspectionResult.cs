﻿using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class ApplicationWorksheetFunctionInspectionResult : InspectionResultBase
    {
        public ApplicationWorksheetFunctionInspectionResult(IInspection inspection, QualifiedSelection qualifiedSelection, IdentifierReference reference)
            : base(inspection, qualifiedSelection.QualifiedName, reference.Context, reference.Declaration)
        {
            QualifiedSelection = qualifiedSelection;
        }

        public override QualifiedSelection QualifiedSelection { get; }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ApplicationWorksheetFunctionInspectionResultFormat, Target.IdentifierName).Capitalize(); }
        }
    }
}
