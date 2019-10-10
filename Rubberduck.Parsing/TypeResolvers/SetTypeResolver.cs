using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.TypeResolvers
{
    public class SetTypeResolver :  ISetTypeResolver
    {
        public const string NotAnObject = "NotAnObject";


        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SetTypeResolver(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public Declaration SetTypeDeclaration(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule)
        {
            if (expression == null)
            {
                return null;
            }

            switch (expression)
            {
                case VBAParser.LExprContext lExpression:
                    return SetTypeDeclaration(lExpression.lExpression(), containingModule);
                case VBAParser.NewExprContext newExpression:
                    return SetTypeDeclaration(newExpression.expression(), containingModule);
                case VBAParser.TypeofexprContext typeOfExpression:
                    return SetTypeDeclaration(typeOfExpression.expression(), containingModule);
                default:
                    return null; //All remaining expression types either have no Set type or there is no declaration for it. 
            }
        }

        private Declaration SetTypeDeclaration(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var setTypeDeterminingDeclaration =
                SetTypeDeterminingDeclarationOfExpression(lExpression, containingModule, finder);
            return SetTypeDeclaration(setTypeDeterminingDeclaration.declaration);
        }

        private Declaration SetTypeDeclaration(Declaration setTypeDeterminingDeclaration)
        {
            return setTypeDeterminingDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? true
                ? setTypeDeterminingDeclaration
                : setTypeDeterminingDeclaration.AsTypeDeclaration;
        }


        public string SetTypeName(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule)
        {
            if (expression == null)
            {
                return null;
            }

            switch (expression)
            {
                case VBAParser.LExprContext lExpression:
                    return SetTypeName(lExpression.lExpression(), containingModule);
                case VBAParser.NewExprContext newExpression:
                    return SetTypeName(newExpression.expression(), containingModule);
                case VBAParser.TypeofexprContext typeOfExpression:
                    return SetTypeName(typeOfExpression.expression(), containingModule);
                case VBAParser.LiteralExprContext literalExpression:
                    return SetTypeName(literalExpression.literalExpression());
                case VBAParser.BuiltInTypeExprContext builtInTypeExpression:
                    return SetTypeName(builtInTypeExpression.builtInType());
                default:
                    return NotAnObject; //All remaining expression types have no Set type. 
            }
        }

        private string SetTypeName(VBAParser.LiteralExpressionContext literalExpression)
        {
            var literalIdentifier = literalExpression.literalIdentifier();

            if (literalIdentifier?.objectLiteralIdentifier() != null)
            {
                return Tokens.Object;
            }

            return NotAnObject;
        }

        private string SetTypeName(VBAParser.BuiltInTypeContext builtInType)
        {
            if (builtInType.OBJECT() != null)
            {
                return Tokens.Object;
            }

            var baseType = builtInType.baseType();

            if (baseType.VARIANT() != null)
            {
                return Tokens.Variant;
            }

            if (baseType.ANY() != null)
            {
                return Tokens.Any;
            }

            return NotAnObject;
        }

        private string SetTypeName(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var setTypeDeterminingDeclaration = SetTypeDeterminingDeclarationOfExpression(lExpression, containingModule, finder);
            return setTypeDeterminingDeclaration.mightHaveSetType
                ? FullObjectTypeName(setTypeDeterminingDeclaration.declaration, lExpression)
                : NotAnObject;
        }

        private static string FullObjectTypeName(Declaration setTypeDeterminingDeclaration, VBAParser.LExpressionContext lExpression)
        {
            if (setTypeDeterminingDeclaration == null)
            {
                return null;
            }

            if (setTypeDeterminingDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                return setTypeDeterminingDeclaration.QualifiedModuleName.ToString();
            }

            if (setTypeDeterminingDeclaration.IsObject || setTypeDeterminingDeclaration.IsObjectArray)
            {
                return setTypeDeterminingDeclaration.FullAsTypeName;
            }

            return setTypeDeterminingDeclaration.AsTypeName == Tokens.Variant
                ? setTypeDeterminingDeclaration.AsTypeName
                : NotAnObject;
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            switch (lExpression)
            {
                case VBAParser.SimpleNameExprContext simpleNameExpression:
                    return SetTypeDeterminingDeclarationOfExpression(simpleNameExpression.identifier(), containingModule, finder);
                case VBAParser.InstanceExprContext instanceExpression:
                    return SetTypeDeterminingDeclarationOfInstance(containingModule, finder);
                case VBAParser.IndexExprContext indexExpression:
                    return SetTypeDeterminingDeclarationOfIndexExpression(indexExpression, containingModule, finder);
                case VBAParser.MemberAccessExprContext memberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(memberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.WithMemberAccessExprContext withMemberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(withMemberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.DictionaryAccessExprContext dictionaryAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(dictionaryAccessExpression.dictionaryAccess(), containingModule, finder);
                case VBAParser.WithDictionaryAccessExprContext withDictionaryAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(withDictionaryAccessExpression.dictionaryAccess(), containingModule, finder);
                case VBAParser.WhitespaceIndexExprContext whitespaceIndexExpression:
                    return SetTypeDeterminingDeclarationOfIndexExpression(whitespaceIndexExpression, containingModule, finder);
                default:
                    return (null, true);    //We should already cover every case. Return the value indicating that we have no idea.
            }
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfIndexExpression(VBAParser.LExpressionContext indexExpr, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var lExpressionOfIndexExpression = indexExpr is VBAParser.IndexExprContext indexExpression
                ? indexExpression.lExpression()
                : (indexExpr as VBAParser.WhitespaceIndexExprContext)?.lExpression();

            if (lExpressionOfIndexExpression == null)
            {
                throw new NotSupportedException("Called index expression resolution on expression, which is neither a properly built indexExpr nor a properly built whitespaceIndexExpr.");
            }

            //Passing the indexExpr itself is correct. 
            var arrayDeclaration = ResolveIndexExpressionAsArrayAccess(indexExpr, containingModule, finder);
            if (arrayDeclaration != null)
            {
                return (arrayDeclaration, MightHaveSetTypeOnArrayAccess(arrayDeclaration));
            }

            var declaration = ResolveIndexExpressionAsDefaultMemberAccess(lExpressionOfIndexExpression, containingModule, finder)
                ?? ResolveIndexExpressionAsMethod(lExpressionOfIndexExpression, containingModule, finder);
                              
            return (declaration, MightHaveSetType(declaration));
        }

        private Declaration ResolveIndexExpressionAsMethod(VBAParser.LExpressionContext lExpressionOfIndexExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            //For functions and properties, the identifier will be at the end of the lExpression.
            var qualifiedSelection = new QualifiedSelection(containingModule, lExpressionOfIndexExpression.GetSelection().Collapse());
            var candidate = finder
                .ContainingIdentifierReferences(qualifiedSelection)
                .LastOrDefault()
                ?.Declaration;
            return candidate?.DeclarationType.HasFlag(DeclarationType.Member) ?? false
                ? candidate
                : null;
        }

        private Declaration ResolveIndexExpressionAsDefaultMemberAccess(VBAParser.LExpressionContext lExpressionOfIndexExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            // A default member access references the entire lExpression.
            // If there are multiple, the references are ordered by recursion depth.
            var qualifiedSelection = new QualifiedSelection(containingModule, lExpressionOfIndexExpression.GetSelection());
            return finder
                .IdentifierReferences(qualifiedSelection)
                .LastOrDefault(reference => reference.IsDefaultMemberAccess)
                ?.Declaration;
        }
        
        //Please note that the lExpression is the (whitespace) index expression itself and not the lExpression it contains. 
        private Declaration ResolveIndexExpressionAsArrayAccess(VBAParser.LExpressionContext actualIndexExpr, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            // An array access references the entire (whitespace)indexExpr.
            var qualifiedSelection = new QualifiedSelection(containingModule, actualIndexExpr.GetSelection());
            return finder
                .IdentifierReferences(qualifiedSelection)
                .LastOrDefault(reference => reference.IsArrayAccess)
                ?.Declaration;
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.IdentifierContext identifier, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var declaration = finder.IdentifierReferences(identifier, containingModule)
                .Select(reference => reference.Declaration)
                .LastOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.UnrestrictedIdentifierContext identifier, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var declaration = finder.IdentifierReferences(identifier, containingModule)
                .Select(reference => reference.Declaration)
                .LastOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.DictionaryAccessContext dictionaryAccess, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var qualifiedSelection = new QualifiedSelection(containingModule, dictionaryAccess.GetSelection());
            var declaration = finder.IdentifierReferences(qualifiedSelection)
                .Select(reference => reference.Declaration)
                .LastOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private static bool MightHaveSetType(Declaration declaration)
        {
            return declaration == null
                   || declaration.IsObject
                   || Tokens.Variant.Equals( declaration.AsTypeName, StringComparison.InvariantCultureIgnoreCase)
                   || declaration.DeclarationType.HasFlag(DeclarationType.ClassModule);
        }

        private static bool MightHaveSetTypeOnArrayAccess(Declaration declaration)
        {
            return declaration == null
                   || declaration.IsObjectArray
                   || Tokens.Variant.Equals(declaration.AsTypeName, StringComparison.InvariantCultureIgnoreCase);
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfInstance(QualifiedModuleName instance, DeclarationFinder finder)
        {
            var classDeclaration = finder.Classes.FirstOrDefault(cls => cls.QualifiedModuleName.Equals(instance));
            return (classDeclaration, true);
        }
    }
}