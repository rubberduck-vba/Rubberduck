using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates procedures that are never invoked from user code.
    /// </summary>
    /// <why>
    /// Unused procedures are dead code that should probably be removed. Note, a procedure may be effectively "not used" in code, but attached to some
    /// Shape object in the host document: in such cases the inspection result should be ignored.
    /// </why>
    /// <remarks>
    /// Not all unused procedures can/should be removed: ignore any inspection results for 
    /// event handler procedures and interface members that Rubberduck isn't recognizing as such.
    /// Public procedures of Standard Modules are not flagged by this inspection regardless of
    /// the presence or absence of user code references.
    /// </remarks>
    /// <example hasResult="true">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Private Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// '@Ignore ProcedureNotUsed
    /// Private Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Macros" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     'a public procedure in a standard module may be a macro 
    ///     'attached to a worksheet Shape or invoked by means other than user code.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// 
    /// Public Sub DoSomethingElse()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub ReferenceOneClass1Procedure()
    ///     Dim target As Class1
    ///     Set target = new Class1
    ///     target.DoSomething
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// 
    /// Public Sub DoSomethingElse()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub ReferenceAllClass1Procedures()
    ///     Dim target As Class1
    ///     Set target = new Class1
    ///     target.DoSomething
    ///     target.DoSomethingElse
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ProcedureNotUsedInspection : DeclarationInspectionBase
    {
        public ProcedureNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, ProcedureTypes)
        {}

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.LibraryProcedure,
            DeclarationType.LibraryFunction,
            DeclarationType.Event
        };

        private static readonly string[] ClassLifeCycleHandlers =
        {
            "Class_Initialize",
            "Class_Terminate"
        };

        private static readonly string[] DocumentEventHandlerPrefixes =
        {
            "Chart_",
            "Worksheet_",
            "Workbook_",
            "Document_",
            "Application_",
            "Session_"
        };

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return !declaration.References
                       .Any(reference => !reference.ParentScoping.Equals(declaration)) // ignore recursive/self-referential calls
                   && !finder.FindEventHandlers().Contains(declaration)
                   && !IsPublicModuleMember(declaration)
                   && !IsClassLifeCycleHandler(declaration)
                   && !(declaration is ModuleBodyElementDeclaration member 
                        && (member.IsInterfaceMember 
                            || member.IsInterfaceImplementation))
                   && !declaration.Annotations
                       .Any(pta => pta.Annotation is ITestAnnotation) 
                   && !IsDocumentEventHandler(declaration);
        }

        /// <remarks>
        /// We cannot determine whether exposed members of standard modules are called or not,
        /// so we assume they are instead of flagging them as "never called".
        /// </remarks>
        private static bool IsPublicModuleMember(Declaration procedure)
        {
            if ((procedure.Accessibility != Accessibility.Implicit
                 && procedure.Accessibility != Accessibility.Public))
            {
                return false;
            }

            var parent = Declaration.GetModuleParent(procedure);
            return parent != null 
                   && parent.DeclarationType.HasFlag(DeclarationType.ProceduralModule);
        }

        private static bool IsClassLifeCycleHandler(Declaration procedure)
        {
            if (!ClassLifeCycleHandlers.Contains(procedure.IdentifierName))
            {
                return false;
            }

            var parent = Declaration.GetModuleParent(procedure);
            return parent != null 
                   && parent.DeclarationType.HasFlag(DeclarationType.ClassModule);
        }

        private static bool IsDocumentEventHandler(Declaration declaration)
        {
            var declarationName = declaration.IdentifierName;
            return DocumentEventHandlerPrefixes
                .Any(prefix => declarationName.StartsWith(prefix));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.IdentifierNotUsedInspection, 
                declarationType, 
                declarationName);
        }
    }
}
