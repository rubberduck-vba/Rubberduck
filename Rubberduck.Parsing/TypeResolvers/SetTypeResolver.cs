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
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SetTypeResolver(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public Declaration SetTypeDeclaration(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var setTypeDeterminingDeclaration = SetTypeDeterminingDeclarationOfExpression(expression, containingModule, finder);
            return SetTypeDeclaration(setTypeDeterminingDeclaration);
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
            return FullObjectTypeName(setTypeDeterminingDeclaration);
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
                : null;
        }

        private Declaration SetTypeDeterminingDeclarationOfExpression(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            switch (expression)
            {
                case VBAParser.LExprContext lExpression:
                    return SetTypeDeterminingDeclarationOfExpression(lExpression.lExpression(), containingModule, finder);
                case VBAParser.NewExprContext newExpression:
                    return null; //Not implemented yet, but it fails inspection tests on the wrong set assignment if we throw here.
                    //throw new NotImplementedException();
                case VBAParser.TypeofexprContext typeOfExpression:
                    throw new NotImplementedException();
                case VBAParser.LiteralExprContext literalExpression:
                    throw new NotImplementedException();
                case VBAParser.BuiltInTypeExprContext builtInTypeExpression:
                    throw new NotImplementedException();
                default:
                    return null;
            }
        }

        private Declaration SetTypeDeterminingDeclarationOfExpression(VBAParser.LExpressionContext lExpression, QualifiedModuleName containingModule, DeclarationFinder finder)
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
                    throw new NotImplementedException();
                case VBAParser.WithMemberAccessExprContext withMemberAccessExpression:
                    throw new NotImplementedException();
                case VBAParser.DictionaryAccessExprContext dictionaryAccessExpression:
                    throw new NotImplementedException();
                case VBAParser.WithDictionaryAccessExprContext withDictionaryAccessExpression:
                    throw new NotImplementedException();
                case VBAParser.WhitespaceIndexExprContext whitespaceIndexExpression:
                    throw new NotImplementedException();
                default:
                    return null;
            }
        }

        private Declaration SetTypeDeterminingDeclarationOfExpression(VBAParser.IdentifierContext identifier, QualifiedModuleName containingModule, DeclarationFinder finder)
        {
            return finder.IdentifierReferences(identifier, containingModule)
                .Select(reference => reference.Declaration)
                .FirstOrDefault(declaration => declaration.IsObject || declaration.AsTypeName == Tokens.Variant);
        }

        private Declaration SetTypeDeterminingDeclarationOfInstance(QualifiedModuleName instance, DeclarationFinder finder)
        {
            return finder.Classes.FirstOrDefault(cls => cls.QualifiedModuleName.Equals(instance));
        }
    }
}