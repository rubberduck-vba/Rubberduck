using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using System.ComponentModel;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public enum ExtractMethodParameterType
    {
        PrivateLocalVariable,
        StaticLocalVariable,
        PrivateModuleField,
        PublicModuleField,
        ByRefParameter,
        ByValParameter
    }

    public enum ExtractMethodParameterFormat
    {
        DimOrParameterDeclaration,
        DimOrParameterDeclarationWithAccessibility,
        ReturnDeclaration
    }

    public class ExtractMethodParameter : INotifyPropertyChanged
    {
        private const string ArrayDim = "()";
        public static readonly string NoneLabel = RubberduckUI.ExtractMethod_OutputNone;

        public event PropertyChangedEventHandler PropertyChanged;

        public ExtractMethodParameter(string typeName, ExtractMethodParameterType parameterType, string name, bool isArray)
        {
            Name = name ?? NoneLabel;
            TypeName = typeName;
            ParameterType = parameterType;
            IsArray = isArray;
        }

        public string Name { get; set; }

        public string TypeName { get; set; }
        
        private ExtractMethodParameterType _parameterType;
        public ExtractMethodParameterType ParameterType
        {
            get => _parameterType;
            set
            {
                _parameterType = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ParameterType)));
            }
        }

        public bool IsArray { get; set; }

        public string ToString(ExtractMethodParameterFormat format)
        {
            switch (format)
            {
                case ExtractMethodParameterFormat.DimOrParameterDeclaration:
                    return string.Concat(Name, IsArray ? ArrayDim : string.Empty, " ", Tokens.As, " ", TypeName);
                case ExtractMethodParameterFormat.ReturnDeclaration:
                    return string.Concat(TypeName, IsArray ? ArrayDim : string.Empty);
                case ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility:
                    return ToString();
                default:
                    return null;
            }
        }

        public override string ToString()
        {
            string accessibility;
            switch (ParameterType)
            {
                case ExtractMethodParameterType.PrivateLocalVariable:
                    accessibility = Tokens.Dim;
                    break;
                case ExtractMethodParameterType.StaticLocalVariable:
                    accessibility = Tokens.Static;
                    break;
                case ExtractMethodParameterType.PrivateModuleField:
                    accessibility = Tokens.Private;
                    break;
                case ExtractMethodParameterType.PublicModuleField:
                    accessibility = Tokens.Public;
                    break;
                case ExtractMethodParameterType.ByRefParameter:
                    accessibility = string.Empty;
                    break;
                case ExtractMethodParameterType.ByValParameter:
                    accessibility = Tokens.ByVal;
                    break;
                default:
                    accessibility = string.Empty;
                    break;
            }
            if (!string.IsNullOrWhiteSpace(accessibility))
            {
                accessibility += " ";
            }
            return string.Concat(accessibility, Name, IsArray ? ArrayDim : string.Empty, " ", Tokens.As, " ", TypeName);
        }

        public static ExtractMethodParameter None => new ExtractMethodParameter(string.Empty,
            ExtractMethodParameterType.PrivateLocalVariable,
            RubberduckUI.ExtractMethod_NoneSelected, false);

        public static Dictionary<ExtractMethodParameterType, string> ParameterTypes
        {
            get
            {
                var dict = new Dictionary<ExtractMethodParameterType, string>
                {
                    {
                        ExtractMethodParameterType.PrivateLocalVariable,
                        RubberduckUI.ExtractParameterNewType_PrivateLocalVariable
                    },
                    {
                        ExtractMethodParameterType.StaticLocalVariable,
                        RubberduckUI.ExtractParameterNewType_StaticLocalVariable
                    },
                    {
                        ExtractMethodParameterType.PrivateModuleField,
                        RubberduckUI.ExtractParameterNewType_PrivateModuleField
                    },
                    {
                        ExtractMethodParameterType.PublicModuleField,
                        RubberduckUI.ExtractParameterNewType_PublicModuleField
                    },
                    {
                        ExtractMethodParameterType.ByRefParameter,
                        RubberduckUI.ExtractParameterNewType_ByRefParameter
                    },
                    {
                        ExtractMethodParameterType.ByValParameter,
                        RubberduckUI.ExtractParameterNewType_ByValParameter
                    }
                };

                return dict;
            }
        }

        public static bool operator ==(ExtractMethodParameter left, ExtractMethodParameter right)
        {
            return left?.TypeName == right?.TypeName && left?.Name == right?.Name;
        }

        public static bool operator !=(ExtractMethodParameter left, ExtractMethodParameter right)
        {
            return !(left?.TypeName == right?.TypeName && left?.Name == right?.Name);
        }
    }
}
