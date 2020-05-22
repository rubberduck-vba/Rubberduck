using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.TypeResolvers;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates assignments to object variables for which the RHS does not have a compatible declared type. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether different object types are compatible. Instead there is a runtime error whenever the types are incompatible.
    /// </why>
    /// <example hasresult="true">
    /// <module name="IInterface" type="Class Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    ///'No Implements IInterface
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim intrfc As IInterface
    ///
    ///     Set cls = New Class1
    ///     Set intrfc = cls 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="IInterface" type="Class Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Implements IInterface
    /// 
    /// Private Sub IInterface_DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim intrfc As IInterface
    ///
    ///     Set cls = New Class1
    ///     Set intrfc = cls 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class SetAssignmentWithIncompatibleObjectTypeInspection : IdentifierReferenceInspectionBase<string>
    {
        private readonly ISetTypeResolver _setTypeResolver;

        public SetAssignmentWithIncompatibleObjectTypeInspection(IDeclarationFinderProvider declarationFinderProvider, ISetTypeResolver setTypeResolver)
            : base(declarationFinderProvider)
        {
            _setTypeResolver = setTypeResolver;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override (bool isResult, string properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            if (!ToBeConsidered(reference))
            {
                return (false, null);
            }

            var assignedTypeName = AssignedTypeName(reference, finder);

            if(assignedTypeName == null || SetAssignmentPossiblyLegal(reference.Declaration, assignedTypeName))
            {
                return (false, null);
            }

            return (true, assignedTypeName);
        }

        private static bool ToBeConsidered(IdentifierReference reference)
        {
            if (reference == null || !reference.IsSetAssignment)
            {
                return false;
            }

            var declaration = reference.Declaration;
            return declaration?.AsTypeDeclaration != null 
                   && declaration.IsObject;
        }

        private string AssignedTypeName(IdentifierReference setAssignment, DeclarationFinder finder)
        {
            return SetTypeNameOfExpression(RHS(setAssignment), setAssignment.QualifiedModuleName, finder);
        }

        private VBAParser.ExpressionContext RHS(IdentifierReference setAssignment)
        {
            return setAssignment.Context.GetAncestor<VBAParser.SetStmtContext>().expression();
        }

        private string SetTypeNameOfExpression(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            return _setTypeResolver.SetTypeName(expression, containingModule);
        }

        private bool SetAssignmentPossiblyLegal(Declaration declaration, string assignedTypeName)
        {
            return assignedTypeName == declaration.FullAsTypeName
                || assignedTypeName == Tokens.Variant
                || assignedTypeName == Tokens.Object
                || HasBaseType(declaration, assignedTypeName)
                || HasSubType(declaration, assignedTypeName)
                || assignedTypeName.EndsWith("stdole.IUnknown")
                || declaration.FullAsTypeName.EndsWith("stdole.IUnknown");
        }

        private bool HasBaseType(Declaration declaration, string typeName)
        {
            var ownType = declaration.AsTypeDeclaration;
            if (ownType == null || !(ownType is ClassModuleDeclaration classType))
            {
                return false;
            }

            return classType.Subtypes.Select(subtype => subtype.QualifiedModuleName.ToString()).Contains(typeName);
        }

        private bool HasSubType(Declaration declaration, string typeName)
        {
            var ownType = declaration.AsTypeDeclaration;
            if (ownType == null || !(ownType is ClassModuleDeclaration classType))
            {
                return false;
            }

            return classType.Supertypes.Select(supertype => supertype.QualifiedModuleName.ToString()).Contains(typeName);
        }

        protected override string ResultDescription(IdentifierReference reference, string assignedTypeName)
        {
            var declarationName = reference.Declaration.IdentifierName;
            var variableTypeName = reference.Declaration.FullAsTypeName;
            return string.Format(InspectionResults.SetAssignmentWithIncompatibleObjectTypeInspection, declarationName, variableTypeName, assignedTypeName);
        }
    }
}
