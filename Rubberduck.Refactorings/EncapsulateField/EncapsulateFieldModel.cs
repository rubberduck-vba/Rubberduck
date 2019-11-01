using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateField
    {
        Declaration TargetDeclaration { get; set; }
        string PropertyName { get; }
        string ParameterName { get; }
        bool ImplementLetSetterType { get;}
        bool ImplementSetSetterType { get;}

        bool CanImplementLet { get; }
        bool CanImplementSet { get; }
    }

    public interface IEncapsulateUDTMember : IEncapsulateField
    {
        bool Encapsulate { set; get; }
    }

    public interface IEncapsulateFieldModel
    {
        IEnumerable<IEncapsulateField> EncapsulationTargets { get; }
    }

    public struct EncapsulateFieldConfig : IEncapsulateField
    {
        public EncapsulateFieldConfig(Declaration target)
        {
            TargetDeclaration = target;
            PropertyName = target.IdentifierName;
            ParameterName = "value";
            ImplementLetSetterType = true;
            ImplementSetSetterType = true;
            CanImplementLet = true;
            CanImplementSet = true;
            Encapsulate = true;
            AssignSetterAndLetterAvailability(target, out var canImplSet, out var canImplLet);
            CanImplementLet = canImplLet;
            CanImplementSet = canImplSet;
        }

        public EncapsulateFieldConfig(IEncapsulateField config)
        {
            TargetDeclaration = config.TargetDeclaration;
            PropertyName = config.PropertyName;
            ParameterName = config.ParameterName;
            ImplementLetSetterType = config.ImplementLetSetterType;
            ImplementSetSetterType = config.ImplementSetSetterType;
            CanImplementLet = config.CanImplementLet;
            CanImplementSet = config.CanImplementSet;
            Encapsulate = config is IEncapsulateUDTMember udt ? udt.Encapsulate : true;
        }

        public Declaration TargetDeclaration { get; set; }
        public string PropertyName { get; set; }
        public string ParameterName { get; set; }
        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }

        public bool CanImplementLet { get; private set; }
        public bool CanImplementSet { get; private set; }
        public bool Encapsulate { get; set; }

        private static void AssignSetterAndLetterAvailability(Declaration targetDeclaration, out bool canImplSet, out bool canImplLet)
        {
            var isVariant = targetDeclaration.AsTypeName.Equals(Tokens.Variant);
            var isValueType = !isVariant && (SymbolList.ValueTypes.Contains(targetDeclaration.AsTypeName) ||
                                             targetDeclaration.DeclarationType == DeclarationType.Enumeration);

            canImplSet = !isValueType;
            canImplLet = isValueType || isVariant;
        }
    }

    public class EncapsulateFieldModel : IRefactoringModel, IEncapsulateFieldModel, IEncapsulateField
    {
        public EncapsulateFieldModel(Declaration target)
        {
            AddTarget(target);
        }

        private readonly object _targetsLock = new object();
        private Dictionary<Declaration, IEncapsulateField> _targets = new Dictionary<Declaration, IEncapsulateField>();

        public void AddTarget(Declaration target)
        {
            lock (_targetsLock)
            {
                _targets.Add(target, new EncapsulateFieldConfig(target));
            }
        }

        public void AddTarget(IEncapsulateField encapsulateFieldConfig)
        {
            lock (_targetsLock)
            {
                if (_targets.ContainsKey(encapsulateFieldConfig.TargetDeclaration))
                {
                    _targets[encapsulateFieldConfig.TargetDeclaration] = encapsulateFieldConfig;
                    return;
                }
                _targets.Add(encapsulateFieldConfig.TargetDeclaration, encapsulateFieldConfig);
            }
        }

        private IEncapsulateField DefaultTarget
        {
            get
            {
                lock (_targetsLock)
                {
                    return _targets.Values.ElementAt(0);
                };
            }
        }


        public IEnumerable<IEncapsulateField> EncapsulationTargets 
        {
            get
            {
                var result = Enumerable.Empty<IEncapsulateField>();
                lock (_targetsLock)
                {
                    result = _targets.Values.Select(v => v as IEncapsulateField).ToList();
                };
                return result;
            }
        }

        public Declaration TargetDeclaration
        {
            get => DefaultTarget.TargetDeclaration;
            set
            {
                AddTarget(value);
            }
        }

        public string PropertyName
        {
            get => DefaultTarget.PropertyName;
            set
            {
                var config = new EncapsulateFieldConfig(DefaultTarget)
                {
                    PropertyName = value
                };
                AddTarget(config);
            }
        }

        public string ParameterName
        {
            get => DefaultTarget.ParameterName;
            set
            {
                var config = new EncapsulateFieldConfig(DefaultTarget)
                {
                    ParameterName = value
                };
                AddTarget(config);
            }
        }

        public bool ImplementLetSetterType
        {
            get => DefaultTarget.ImplementLetSetterType;
            set
            {
                var config = new EncapsulateFieldConfig(DefaultTarget)
                {
                    ImplementLetSetterType = value
                };
                AddTarget(config);
            }
        }

        public bool ImplementSetSetterType
        {
            get => DefaultTarget.ImplementSetSetterType;
            set
            {
                var config = new EncapsulateFieldConfig(DefaultTarget)
                {
                    ImplementSetSetterType = value
                };
                AddTarget(config);
            }
        }

        public bool CanImplementLet
            => DefaultTarget.CanImplementLet;

        public bool CanImplementSet
            => DefaultTarget.CanImplementSet;
    }
}
