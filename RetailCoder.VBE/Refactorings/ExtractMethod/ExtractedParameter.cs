using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using System.ComponentModel;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public enum ExtractParameterNewType
    {
        PrivateLocalVariable,
        StaticLocalVariable,
        PrivateModuleField,
        PublicModuleField,
        ByRefParameter,
        ByValParameter
    }

    public class ExtractParameterNewTypeDescription
    {
        public IEnumerable Text(ExtractParameterNewType extractParameterType)
        {
            switch (extractParameterType)
            {
                case ExtractParameterNewType.PrivateLocalVariable:
                    return RubberduckUI.ExtractParameterNewType_PrivateLocalVariable;
                case ExtractParameterNewType.StaticLocalVariable:
                    return RubberduckUI.ExtractParameterNewType_StaticLocalVariable;
                case ExtractParameterNewType.PrivateModuleField:
                    return RubberduckUI.ExtractParameterNewType_PrivateModuleField;
                case ExtractParameterNewType.PublicModuleField:
                    return RubberduckUI.ExtractParameterNewType_PublicModuleField;
                case ExtractParameterNewType.ByRefParameter:
                    return RubberduckUI.ExtractParameterNewType_ByRefParameter;
                case ExtractParameterNewType.ByValParameter:
                    return RubberduckUI.ExtractParameterNewType_ByValParameter;
                default:
                    throw new InvalidOperationException("Invalid value given for extractParameterType");
            }
        }
    }

    public class ExtractedParameter : INotifyPropertyChanged
    {
        public static readonly string None = RubberduckUI.ExtractMethod_OutputNone;

        public event PropertyChangedEventHandler PropertyChanged;

        public ExtractedParameter(string typeName, ExtractParameterNewType parameterType, string name = null)
        {
            Name = name ?? None;
            TypeName = typeName;
            ParameterType = parameterType;
        }

        public string Name { get; set; }

        public string TypeName { get; set; }

        private ExtractParameterNewType _parameterType;
        public ExtractParameterNewType ParameterType
        {
            get => _parameterType;
            set
            {
                _parameterType = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ParameterType)));
            }
        }

        public override string ToString()
        {
            return ParameterType.ToString() + ' ' + Name + ' ' + Tokens.As + ' ' + TypeName;
        }

        public static Dictionary<ExtractParameterNewType, string> ParameterNewTypes
        {
            get
            {
                var dict = new Dictionary<ExtractParameterNewType, string>
                {
                    {
                        ExtractParameterNewType.PrivateLocalVariable,
                        RubberduckUI.ExtractParameterNewType_PrivateLocalVariable
                    },
                    {
                        ExtractParameterNewType.StaticLocalVariable,
                        RubberduckUI.ExtractParameterNewType_StaticLocalVariable
                    },
                    {
                        ExtractParameterNewType.PrivateModuleField,
                        RubberduckUI.ExtractParameterNewType_PrivateModuleField
                    },
                    {ExtractParameterNewType.PublicModuleField, RubberduckUI.ExtractParameterNewType_PublicModuleField},
                    {ExtractParameterNewType.ByRefParameter, RubberduckUI.ExtractParameterNewType_ByRefParameter},
                    {ExtractParameterNewType.ByValParameter, RubberduckUI.ExtractParameterNewType_ByValParameter}
                };

                return dict;
            }
        }
    }
}
