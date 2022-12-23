using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using System.ComponentModel;
using System.Collections.ObjectModel;

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

        public ExtractMethodParameter(string typeName, ExtractMethodParameterType parameterType, string name, bool isArray, bool isObject, bool canReturn)
        {
            Name = name ?? NoneLabel;
            TypeName = typeName;
            ParameterType = parameterType;
            IsArray = isArray;
            CanReturn = canReturn;
            IsObject = isObject;
        }

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


        public static ExtractMethodParameter None => new ExtractMethodParameter(string.Empty,
            ExtractMethodParameterType.ByValParameter,
            RefactoringsUI.ExtractMethod_NoneSelected, false, false, false);

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
                   left?.TypeName == right?.TypeName &&
                   left?.Name == right?.Name &&
                   left?.IsArray == right?.IsArray &&
                   left?.CanReturn == right?.CanReturn &&
                   left?.IsObject == right?.IsObject;
        }

        public static bool operator !=(ExtractMethodParameter left, ExtractMethodParameter right)
        {
            return !(left?.ParameterType == right?.ParameterType &&
                     left?.TypeName == right?.TypeName &&
                     left?.Name == right?.Name &&
                     left?.IsArray == right?.IsArray &&
                     left?.CanReturn == right?.CanReturn &&
                     left?.IsObject == right?.IsObject);
        }

        public override bool Equals(object obj)
        {
            return obj is ExtractMethodParameter parameter &&
                   Name == parameter.Name &&
                   TypeName == parameter.TypeName &&
                   CanReturn == parameter.CanReturn &&
                   ParameterType == parameter.ParameterType &&
                   IsArray == parameter.IsArray &&
                   IsObject == parameter.IsObject;
        }

        public override int GetHashCode()
        {
            int hashCode = 1661774273;
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(TypeName);
            hashCode = (hashCode * -1521134295) + CanReturn.GetHashCode();
            hashCode = (hashCode * -1521134295) + ParameterType.GetHashCode();
            hashCode = (hashCode * -1521134295) + IsArray.GetHashCode();
            hashCode = (hashCode * -1521134295) + IsObject.GetHashCode();
            return hashCode;
        }
    }
}
