﻿using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates places in which a procedure needs to be called but an object variables has been provided that does not have a suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Sub Foo()
    /// 'No default member attribute
    /// End Sub
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     cls 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Sub Foo()
    /// Attribute Foo.UserMemId = 0
    /// End Sub
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     cls 
    /// End Sub
    /// ]]>
    /// </example>
    public class ProcedureRequiredInspection : IdentifierReferenceInspectionBase
    {
        public ProcedureRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.FailedProcedureCoercions(module);
        }

        protected override bool IsResultReference(IdentifierReference failedCoercion)
        {
            // return true because no special ignore checking is required
            return true;
        }

        protected override string ResultDescription(IdentifierReference failedCoercion)
        {
            var expression = failedCoercion.IdentifierName;
            var typeName = failedCoercion.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.ProcedureRequiredInspection, expression, typeName);
        }
    }
}