using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using System.ComponentModel;
using System.Collections.ObjectModel;
using Rubberduck.Parsing.Symbols;
using System.Xml.Linq;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public enum ExtractMethodParameterType
    {
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
        public static readonly string NoneLabel = RefactoringsUI.ExtractMethod_NoneSelected;

        public event PropertyChangedEventHandler PropertyChanged;

        public ExtractMethodParameter(Declaration declaration, ExtractMethodParameterType parameterType, bool canReturn)
        {
            ParameterType = parameterType;
            CanReturn = canReturn;
            Declaration = declaration;
            if (declaration == null)
            {
                Name = NoneLabel;
                TypeName = string.Empty;
                IsArray = false;
                IsObject = false;
            }
            else
            {
                Name = declaration.IdentifierName;
                TypeName = declaration.AsTypeNameWithoutArrayDesignator;
                IsArray = declaration.IsArray;
                IsObject = declaration.IsObject;
            }
        }

        public Declaration Declaration { get; }
        public string Name { get; set; }

        public string TypeName { get; set; }

        public string ParameterTypeText
        {
            get
            {
                switch (ParameterType)
                {
                    case ExtractMethodParameterType.ByRefParameter:
                        return Tokens.ByRef;
                    case ExtractMethodParameterType.ByValParameter:
                        return Tokens.ByVal;
                    default:
                        return string.Empty;
                }
            }
        }
        public bool CanReturn { get; set; }

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

        public bool IsObject { get; set; }

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
                case ExtractMethodParameterType.ByRefParameter:
                    accessibility = Tokens.ByRef;
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


        public static ExtractMethodParameter None => new ExtractMethodParameter(null, ExtractMethodParameterType.ByValParameter, false);

        public static ObservableCollection<ExtractMethodParameterType> AllowableTypes = new ObservableCollection<ExtractMethodParameterType>
            {
                ExtractMethodParameterType.ByRefParameter,
                ExtractMethodParameterType.ByValParameter
            };

        public static Dictionary<ExtractMethodParameterType, string> ParameterTypes
        {
            get
            {
                var dict = new Dictionary<ExtractMethodParameterType, string>
                {
                    {
                        ExtractMethodParameterType.ByRefParameter,
                        RefactoringsUI.ExtractParameterNewType_ByRefParameter
                    },
                    {
                        ExtractMethodParameterType.ByValParameter,
                        RefactoringsUI.ExtractParameterNewType_ByValParameter
                    }
                };

                return dict;
            }
        }

        public static bool operator ==(ExtractMethodParameter left, ExtractMethodParameter right)
        {
            return left?.ParameterType == right?.ParameterType &&
                   left?.CanReturn == right?.CanReturn &&
                   left?.Declaration == right?.Declaration;
        }

        public static bool operator !=(ExtractMethodParameter left, ExtractMethodParameter right)
        {
            return !(left?.ParameterType == right?.ParameterType &&
                     left?.CanReturn == right?.CanReturn &&
                     left?.Declaration == right?.Declaration);
        }

        public override bool Equals(object obj)
        {
            return obj is ExtractMethodParameter parameter &&
                   CanReturn == parameter.CanReturn &&
                   ParameterType == parameter.ParameterType &&
                   Declaration == parameter.Declaration;
        }

        public override int GetHashCode()
        {
            int hashCode = 1661774273;
            hashCode = (hashCode * -1521134295) + CanReturn.GetHashCode();
            hashCode = (hashCode * -1521134295) + ParameterType.GetHashCode();
            hashCode = (hashCode * -1521134295) + (Declaration == null ? 0 : Declaration.GetHashCode());
            return hashCode;
        }
    }
}
