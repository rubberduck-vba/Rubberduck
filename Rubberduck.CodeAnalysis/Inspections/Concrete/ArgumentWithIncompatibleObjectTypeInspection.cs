using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
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
    /// Locates arguments passed to functions or procedures for object parameters which the do not have a compatible declared type. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether different object types are compatible. Instead there is a runtime error whenever the types are incompatible.
    /// </why>
    /// <example hasresult="true">
    /// <module name="Interface" type="Class Module">
    /// <![CDATA[
    /// IInterface:
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// 'No Implements IInterface
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     Foo cls 
    /// End Sub
    ///
    /// Public Sub Foo(cls As IInterface)
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="Interface" type="Class Module">
    /// <![CDATA[
    /// IInterface:
    ///
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Implements IInterface
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     Foo cls 
    /// End Sub
    ///
    /// Public Sub Foo(cls As IInterface)
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class ArgumentWithIncompatibleObjectTypeInspection : ArgumentReferenceInspectionFromDeclarationsBase<string>
    {
        private readonly ISetTypeResolver _setTypeResolver;

        public ArgumentWithIncompatibleObjectTypeInspection(IDeclarationFinderProvider declarationFinderProvider, ISetTypeResolver setTypeResolver)
            : base(declarationFinderProvider)
        {
            _setTypeResolver = setTypeResolver;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return finder.DeclarationsWithType(DeclarationType.Parameter)
                .Where(ToBeConsidered);
        }

        private static bool ToBeConsidered(Declaration declaration)
        {
            return declaration?.AsTypeDeclaration != null
                   && declaration.IsObject;
        }

        protected override (bool isResult, string properties) IsUnsuitableArgumentWithAdditionalProperties(ArgumentReference reference, DeclarationFinder finder)
        {
            var argumentSetTypeName = ArgumentSetTypeName(reference, finder);

            if (argumentSetTypeName == null || ArgumentPossiblyLegal(reference.Declaration, argumentSetTypeName))
            {
                return (false, null);
            }

            return (true, argumentSetTypeName);
        }

        private string ArgumentSetTypeName(IdentifierReference argumentReference, DeclarationFinder finder)
        {
            var argumentExpression = argumentReference.Context as VBAParser.ExpressionContext;
            return SetTypeNameOfExpression(argumentExpression, argumentReference.QualifiedModuleName, finder);
        }

        private string SetTypeNameOfExpression(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            return _setTypeResolver.SetTypeName(expression, containingModule);
        }

        private bool ArgumentPossiblyLegal(Declaration parameterDeclaration , string assignedTypeName)
        {
            return assignedTypeName == parameterDeclaration.FullAsTypeName
                || assignedTypeName == Tokens.Variant
                || assignedTypeName == Tokens.Object
                || HasBaseType(parameterDeclaration, assignedTypeName)
                || HasSubType(parameterDeclaration, assignedTypeName)
                || assignedTypeName.EndsWith("stdole.IUnknown")
                || parameterDeclaration.FullAsTypeName.EndsWith("stdole.IUnknown");
        }

        private static bool HasBaseType(Declaration declaration, string typeName)
        {
            var ownType = declaration.AsTypeDeclaration;
            if (ownType == null || !(ownType is ClassModuleDeclaration classType))
            {
                return false;
            }

            return classType.Subtypes.Select(subtype => subtype.QualifiedModuleName.ToString()).Contains(typeName);
        }

        private static bool HasSubType(Declaration declaration, string typeName)
        {
            var ownType = declaration.AsTypeDeclaration;
            if (ownType == null || !(ownType is ClassModuleDeclaration classType))
            {
                return false;
            }

            return classType.Supertypes.Select(supertype => supertype.QualifiedModuleName.ToString()).Contains(typeName);
        }

        protected override string ResultDescription(IdentifierReference reference, string argumentTypeName)
        {
            var parameterName = reference.Declaration.IdentifierName;
            var parameterTypeName = reference.Declaration.FullAsTypeName;
            var argumentExpression = reference.Context.GetText();
            return string.Format(InspectionResults.SetAssignmentWithIncompatibleObjectTypeInspection, parameterName, parameterTypeName, argumentExpression, argumentTypeName);
        }
    }
}
