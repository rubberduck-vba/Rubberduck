using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberResources
    {
        public static string Class_Initialize => "Class_Initialize";
        public static string Class_Terminate => "Class_Terminate";
        public static string UserForm => "UserForm";
        public static string OptionExplicit => $"{Tokens.Option} {Tokens.Explicit}";

        public static string Caption => "MoveMember";
        public static string DefaultErrorMessageFormat => "Unable to Move Member: {0}";
        public static string InvalidMoveDefinition => "Incomplete Move definition: Code element(s) and destination module must be defined";
        public static string VBALanguageSpecificationViolation => "The defined Move would result in a VBA Language Specification violation and generate uncompilable code";
        public static string ApplicableStrategyNotFound => "Applicable move strategy not found";
        public static string Prefix_Variable => "xxx_";
        public static string Prefix_Parameter => "x";  //"arg_";
        public static string Prefix_ClassInstantiationProcedure => "X_"; // "Create__";
        public static string UnsupportedMoveExceptionFormat => "Unable to Move Member: {1}({0})";


        //TODO: Make this a DeclarationExtension bool MatchesLifeCycleHandlerSignature()
        public static bool IsOrNamedLikeALifeCycleHandler(Declaration member)
            => member.IdentifierName.Equals(Class_Initialize) || member.IdentifierName.Equals(Class_Terminate);
    }
}
