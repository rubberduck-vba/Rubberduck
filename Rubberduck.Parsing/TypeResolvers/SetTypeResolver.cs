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
            var finder = _declarationFinderProvider.DeclarationFinder;
            var setTypeDeterminingDeclaration = SetTypeDeterminingDeclarationOfExpression(expression, containingModule, finder);
            return setTypeDeterminingDeclaration.mightHaveSetType
                ? SetTypeDeclaration(setTypeDeterminingDeclaration.declaration)
                : null;
        }

        private Declaration SetTypeDeclaration(Declaration setTypeDeterminingDeclaration)
        {
            return setTypeDeterminingDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? true
                ? setTypeDeterminingDeclaration
                : setTypeDeterminingDeclaration.AsTypeDeclaration;
        }


        public string SetTypeName(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var setTypeDeterminingDeclaration = SetTypeDeterminingDeclarationOfExpression(expression, containingModule, finder);
            return setTypeDeterminingDeclaration.mightHaveSetType
                ? FullObjectTypeName(setTypeDeterminingDeclaration.declaration)
                : NotAnObject;
        }

        private string FullObjectTypeName(Declaration setTypeDeterminingDeclaration)
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

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfExpression(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            switch (expression)
            {
                case VBAParser.LExprContext lExpression:
                    return SetTypeDeterminingDeclarationOfExpression(lExpression.lExpression(), containingModule, finder);
                case VBAParser.NewExprContext newExpression:
                    return (null, true); //Not implemented yet, but it fails inspection tests on the wrong set assignment if we throw here.
                    //throw new NotImplementedException();
                case VBAParser.TypeofexprContext typeOfExpression:
                    throw new NotImplementedException();
                case VBAParser.LiteralExprContext literalExpression:
                    throw new NotImplementedException();
                case VBAParser.BuiltInTypeExprContext builtInTypeExpression:
                    throw new NotImplementedException();
                default:
                    return (null, false); //All remaining expression types have no Set type. 
            }
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
                    throw new NotImplementedException();
                case VBAParser.MemberAccessExprContext memberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(memberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.WithMemberAccessExprContext withMemberAccessExpression:
                    return SetTypeDeterminingDeclarationOfExpression(withMemberAccessExpression.unrestrictedIdentifier(), containingModule, finder);
                case VBAParser.DictionaryAccessExprContext dictionaryAccessExpression:
                    throw new NotImplementedException();
                case VBAParser.WithDictionaryAccessExprContext withDictionaryAccessExpression:
                    throw new NotImplementedException();
                case VBAParser.WhitespaceIndexExprContext whitespaceIndexExpression:
                    throw new NotImplementedException();
                default:
                    return (null, true);    //We should already cover every case. Return the value indicating that we have no idea.
            }
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

        private static bool MightHaveSetType(Declaration declaration)
        {
            return declaration == null
                   || declaration.IsObject
                   || declaration.AsTypeName == Tokens.Variant
                   || declaration.DeclarationType.HasFlag(DeclarationType.ClassModule);
        }

        private (Declaration declaration, bool mightHaveSetType) SetTypeDeterminingDeclarationOfInstance(QualifiedModuleName instance, DeclarationFinder finder)
        {
            var classDeclaration = finder.Classes.FirstOrDefault(cls => cls.QualifiedModuleName.Equals(instance));
            return (classDeclaration, true);
        }
    }
}