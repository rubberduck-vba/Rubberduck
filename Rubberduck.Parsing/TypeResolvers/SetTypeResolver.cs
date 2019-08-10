using System;
using System.Collections.Generic;
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

            if (setTypeDeterminingDeclaration.IsObject)
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
                    return SetTypeDeterminingDeclarationOfIndexExpression(indexExpression.lExpression(), containingModule, finder);
                case VBAParser.MemberAccessExprContext memberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(memberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.WithMemberAccessExprContext withMemberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(withMemberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.DictionaryAccessExprContext dictionaryAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(dictionaryAccessExpression.dictionaryAccess(), containingModule, finder);
                case VBAParser.WithDictionaryAccessExprContext withDictionaryAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(withDictionaryAccessExpression.dictionaryAccess(), containingModule, finder);
                case VBAParser.WhitespaceIndexExprContext whitespaceIndexExpression:
                    return SetTypeDeterminingDeclarationOfIndexExpression(whitespaceIndexExpression.lExpression(), containingModule, finder);
                default:
                    return (null, true);    //We should already cover every case. Return the value indicating that we have no idea.
            }
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfIndexExpression(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var declaration = ResolveIndexExpressionAsMethod(lExpression, containingModule, finder)
                              ?? ResolveIndexExpressionAsDefaultMemberAccess(lExpression, containingModule, finder);
                
            if (declaration != null)
            {
                return (declaration, MightHaveSetType(declaration));
            }

            return ResolveIndexExpressionAsArrayAccess(lExpression, containingModule, finder);
        }

        private Declaration ResolveIndexExpressionAsMethod(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            //For functions and properties, the identifier will be at the end of the lExpression.
            var qualifiedSelection = new QualifiedSelection(containingModule, lExpression.GetSelection().Collapse());
            var candidate = finder
                .ContainingIdentifierReferences(qualifiedSelection)
                .LastOrDefault()
                ?.Declaration;
            return candidate?.DeclarationType.HasFlag(DeclarationType.Member) ?? false
                ? candidate
                : null;
        }

        private (Declaration declaration, bool mightHaveSetType) ResolveIndexExpressionAsArrayAccess(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var (potentialArrayDeclaration, lExpressionMightHaveSetType) = SetTypeDeterminingDeclarationOfExpression(lExpression, containingModule, finder);

            if (potentialArrayDeclaration == null)
            {
                return (null, lExpressionMightHaveSetType);
            }

            if (!potentialArrayDeclaration.IsArray)
            {
                //This is not an array access. So, we have no idea.
                return (null, true);
            }

            return (potentialArrayDeclaration, MightHaveSetTypeOnArrayAccess(potentialArrayDeclaration));
        }

        private Declaration ResolveIndexExpressionAsDefaultMemberAccess(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            // A default member access references the entire lExpression.
            var qualifiedSelection = new QualifiedSelection(containingModule, lExpression.GetSelection());
            return finder
                .IdentifierReferences(qualifiedSelection)
                .FirstOrDefault(reference => reference.IsDefaultMemberAccess)
                ?.Declaration;
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.IdentifierContext identifier, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var declaration = finder.IdentifierReferences(identifier, containingModule)
                .Select(reference => reference.Declaration)
                .FirstOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.UnrestrictedIdentifierContext identifier, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var declaration = finder.IdentifierReferences(identifier, containingModule)
                .Select(reference => reference.Declaration)
                .FirstOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.DictionaryAccessContext dictionaryAccess, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            var qualifiedSelection = new QualifiedSelection(containingModule, dictionaryAccess.GetSelection());
            var declaration = finder.IdentifierReferences(qualifiedSelection)
                .Select(reference => reference.Declaration)
                .FirstOrDefault();
            return (declaration, MightHaveSetType(declaration));
        }

        private static bool MightHaveSetType(Declaration declaration)
        {
            return declaration == null
                   || declaration.IsObject
                   || declaration.AsTypeName == Tokens.Variant
                   || declaration.DeclarationType.HasFlag(DeclarationType.ClassModule);
        }

        private static bool MightHaveSetTypeOnArrayAccess(Declaration declaration)
        {
            return declaration == null
                || IsObjectArray(declaration)
                || declaration.AsTypeName == Tokens.Variant;
        }

        private static bool IsObjectArray(Declaration declaration)
        {
            if (!declaration.IsArray)
            {
                return false;
            }

            if (declaration.AsTypeName == Tokens.Object ||
                (declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? false))
            {
                return true;
            }

            return false;
        }


        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfInstance(QualifiedModuleName instance, DeclarationFinder finder)
        {
            var classDeclaration = finder.Classes.FirstOrDefault(cls => cls.QualifiedModuleName.Equals(instance));
            return (classDeclaration, true);
        }
    }
}