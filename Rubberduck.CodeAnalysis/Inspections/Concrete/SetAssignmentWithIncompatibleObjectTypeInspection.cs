using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
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
	/// Locates assignments to object variables for which the RHS does not have a compatible declared type. 
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
	///     Dim intrfc As IInterface
	///
	///     Set cls = New Class1
	///     Set intrfc = cls 
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
	///     Dim intrfc As IInterface
	///
	///     Set cls = New Class1
	///     Set intrfc = cls 
	/// End Sub
	/// ]]>
	/// </example>
    public class SetAssignmentWithIncompatibleObjectTypeInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISetTypeResolver _setTypeResolver;

        public SetAssignmentWithIncompatibleObjectTypeInspection(RubberduckParserState state, ISetTypeResolver setTypeResolver)
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

            var setAssignments = finder.AllIdentifierReferences().Where(reference => reference.IsSetAssignment);

            var offendingAssignments = setAssignments
                .Where(ToBeConsidered)
                .Select(setAssignment => SetAssignmentWithAssignedTypeName(setAssignment, finder))
                .Where(setAssignmentWithAssignedTypeName => setAssignmentWithAssignedTypeName.assignedTypeName != null
                                                            && !SetAssignmentPossiblyLegal(setAssignmentWithAssignedTypeName));

            return offendingAssignments
                // Ignoring the Declaration disqualifies all assignments
                .Where(setAssignmentWithAssignedTypeName => !setAssignmentWithAssignedTypeName.setAssignment.Declaration.IsIgnoringInspectionResultFor(AnnotationName))
                .Select(setAssignmentWithAssignedTypeName => InspectionResult(setAssignmentWithAssignedTypeName, _declarationFinderProvider));
        }

        private static bool ToBeConsidered(IdentifierReference reference)
        {
            var declaration = reference.Declaration;
            return declaration != null
                   && declaration.AsTypeDeclaration != null
                   && declaration.IsObject;
        }

        private (IdentifierReference setAssignment, string assignedTypeName) SetAssignmentWithAssignedTypeName(IdentifierReference setAssignment, DeclarationFinder finder)
        {
            return (setAssignment, SetTypeNameOfExpression(RHS(setAssignment), setAssignment.QualifiedModuleName, finder));
        }

        private VBAParser.ExpressionContext RHS(IdentifierReference setAssignment)
        {
            return setAssignment.Context.GetAncestor<VBAParser.SetStmtContext>().expression();
        }

        private string SetTypeNameOfExpression(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            return _setTypeResolver.SetTypeName(expression, containingModule);
        }
        
        private bool SetAssignmentPossiblyLegal((IdentifierReference setAssignment, string assignedTypeName) setAssignmentWithAssignedType)
        {
            var (setAssignment, assignedTypeName) = setAssignmentWithAssignedType;
            
            return SetAssignmentPossiblyLegal(setAssignment.Declaration, assignedTypeName);
        }

        private bool SetAssignmentPossiblyLegal(Declaration declaration, string assignedTypeName)
        {
            return assignedTypeName == declaration.FullAsTypeName
                || assignedTypeName == Tokens.Variant
                || assignedTypeName == Tokens.Object
                || HasBaseType(declaration, assignedTypeName)
                || HasSubType(declaration, assignedTypeName);
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

        private IInspectionResult InspectionResult((IdentifierReference setAssignment, string assignedTypeName) setAssignmentWithAssignedType, IDeclarationFinderProvider declarationFinderProvider)
        {
            var (setAssignment, assignedTypeName) = setAssignmentWithAssignedType;
            return new IdentifierReferenceInspectionResult(this,
                ResultDescription(setAssignment, assignedTypeName),
                declarationFinderProvider,
                setAssignment);
        }

        private string ResultDescription(IdentifierReference setAssignment, string assignedTypeName)
        {
            var declarationName = setAssignment.Declaration.IdentifierName;
            var variableTypeName = setAssignment.Declaration.FullAsTypeName;
            return string.Format(InspectionResults.SetAssignmentWithIncompatibleObjectTypeInspection, declarationName, variableTypeName, assignedTypeName);
        }
    }
}
