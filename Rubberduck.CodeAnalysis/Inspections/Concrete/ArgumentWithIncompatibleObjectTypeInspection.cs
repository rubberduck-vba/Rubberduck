using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
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
	/// <![CDATA[
	/// IInterface:
	///
	/// Public Sub DoSomething()
	/// End Sub
	///
	/// ------------------------------
	/// Class1:
	///
	///'No Implements IInterface
	/// 
	/// Public Sub DoSomething()
	/// End Sub
	///
	/// ------------------------------
	/// Module1:
	/// 
	/// Public Sub DoIt()
	///     Dim cls As Class1
	///     Set cls = New Class1
	///     Foo cls 
	/// End Sub
	///
	/// Public Sub Foo(cls As IInterface)
	/// End Sub
	/// ]]>
	/// </example>
	/// <example hasresult="false">
	/// <![CDATA[
	/// IInterface:
	///
	/// Public Sub DoSomething()
	/// End Sub
	///
	/// ------------------------------
	/// Class1:
	///
	/// Implements IInterface
	/// 
	/// Private Sub IInterface_DoSomething()
	/// End Sub
	///
	/// ------------------------------
	/// Module1:
	/// 
	/// Public Sub DoIt()
	///     Dim cls As Class1
	///     Set cls = New Class1
	///     Foo cls 
	/// End Sub
	///
	/// Public Sub Foo(cls As IInterface)
	/// End Sub
	/// ]]>
	/// </example>
    public class ArgumentWithIncompatibleObjectTypeInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISetTypeResolver _setTypeResolver;

        public ArgumentWithIncompatibleObjectTypeInspection(RubberduckParserState state, ISetTypeResolver setTypeResolver)
            : base(state)
        {
            _declarationFinderProvider = state;
            _setTypeResolver = setTypeResolver;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var strictlyTypedObjectParameters = finder.DeclarationsWithType(DeclarationType.Parameter)
                .Where(ToBeConsidered)
                .OfType<ParameterDeclaration>();

            var offendingArguments = strictlyTypedObjectParameters
                .SelectMany(param => param.ArgumentReferences)
                .Select(argumentReference => ArgumentReferenceWithArgumentTypeName(argumentReference, finder))
                .Where(argumentReferenceWithTypeName =>  argumentReferenceWithTypeName.argumentTypeName != null
                                                         && !ArgumentPossiblyLegal(
                                                             argumentReferenceWithTypeName.argumentReference.Declaration, 
                                                             argumentReferenceWithTypeName.argumentTypeName));

            return offendingArguments
                // Ignoring the Declaration disqualifies all assignments
                .Where(argumentReferenceWithTypeName => !argumentReferenceWithTypeName.Item1.Declaration.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(argumentReference => InspectionResult(argumentReference, _declarationFinderProvider));
        }

        private static bool ToBeConsidered(Declaration declaration)
        {
            return declaration != null
                   && declaration.AsTypeDeclaration != null
                   && declaration.IsObject;
        }

        private (IdentifierReference argumentReference, string argumentTypeName) ArgumentReferenceWithArgumentTypeName(IdentifierReference argumentReference, DeclarationFinder finder)
        {
            return (argumentReference, ArgumentSetTypeName(argumentReference, finder));
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
                || HasSubType(parameterDeclaration, assignedTypeName);
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

        private IInspectionResult InspectionResult((IdentifierReference argumentReference, string argumentTypeName) argumentReferenceWithTypeName, IDeclarationFinderProvider declarationFinderProvider)
        {
            var (argumentReference, argumentTypeName) = argumentReferenceWithTypeName;
            return new IdentifierReferenceInspectionResult(this,
                ResultDescription(argumentReference, argumentTypeName),
                declarationFinderProvider,
                argumentReference);
        }

        private string ResultDescription(IdentifierReference argumentReference, string argumentTypeName)
        {
            var parameterName = argumentReference.Declaration.IdentifierName;
            var parameterTypeName = argumentReference.Declaration.FullAsTypeName;
            var argumentExpression = argumentReference.Context.GetText();
            return string.Format(InspectionResults.SetAssignmentWithIncompatibleObjectTypeInspection, parameterName, parameterTypeName, argumentExpression, argumentTypeName);
        }
    }
}
