using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
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
    public class SetAssignmentWithIncompatibleObjectTypeInspection : IdentifierReferenceInspectionBase
    {
        private readonly ISetTypeResolver _setTypeResolver;

        public SetAssignmentWithIncompatibleObjectTypeInspection(RubberduckParserState state, ISetTypeResolver setTypeResolver)
            : base(state)
        {
            _setTypeResolver = setTypeResolver;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override (bool isResult, object properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
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

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            //No need to implement this since we override IsResultReferenceWithAdditionalProperties.
            throw new System.NotImplementedException();
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

        protected override string ResultDescription(IdentifierReference reference, dynamic properties = null)
        {
            var declarationName = reference.Declaration.IdentifierName;
            var variableTypeName = reference.Declaration.FullAsTypeName;
            var assignedTypeName = (string)properties;
            return string.Format(InspectionResults.SetAssignmentWithIncompatibleObjectTypeInspection, declarationName, variableTypeName, assignedTypeName);
        }
    }
}
