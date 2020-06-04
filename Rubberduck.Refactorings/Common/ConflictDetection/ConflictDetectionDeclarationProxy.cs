using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionModuleDeclarationProxy : IConflictDetectionDeclarationProxy
    {}

    public interface IMutableDeclarationProxy
    {
        string IdentifierName { set; get; }
        DeclarationType DeclarationType { set; get; }
        IEnumerable<IdentifierReference> References { get; }
        Declaration ParentDeclaration { set; get; }
        Accessibility Accessibility { set; get; }
        string ProjectId { set; get; }
    }

    public interface IConflictDetectionDeclarationProxy : IMutableDeclarationProxy
    {
        string TargetModuleName { set; get; }
        ModuleDeclaration TargetModule { set; get; }
        Declaration Prototype { get; }
        IConflictDetectionDeclarationProxy ParentProxy { set; get; }
        QualifiedModuleName? QualifiedModuleName { get; }
        ComponentType ComponentType { set;  get; }
        bool HasStandardModuleParent { get; }
        string ProxyID { get; }
    }

    public class UDTConflictDetectionDeclarationProxy : ConflictDetectionDeclarationProxy, IConflictDetectionDeclarationProxy
    {
        public UDTConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy)
            : base(identifier, declarationType, accessibility, parentProxy.Prototype)
        {
            ParentProxy = parentProxy;
            ProjectId = parentProxy.ProjectId;
            TargetModule = parentProxy.Prototype as ModuleDeclaration;
        }
    }

    public class UDTMemberConflictDetectionDeclarationProxy : ConflictDetectionDeclarationProxy,IConflictDetectionDeclarationProxy
    {
        public UDTMemberConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy)
            : base(identifier, declarationType, accessibility, parentProxy.Prototype)
        {
            ParentProxy = parentProxy;
            ProjectId = parentProxy.ProjectId;
            TargetModule = parentProxy.TargetModule;
        }
    }

    public class EnumConflictDetectionDeclarationProxy : ConflictDetectionDeclarationProxy, IConflictDetectionDeclarationProxy
    {
        public EnumConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy)
            : base(identifier, declarationType, accessibility, parentProxy.Prototype)
        {
            ParentProxy = parentProxy;
            ProjectId = parentProxy.ProjectId;
            TargetModule = parentProxy.Prototype as ModuleDeclaration;
        }
    }

    public class EnumMemberConflictDetectionDeclarationProxy : ConflictDetectionDeclarationProxy, IConflictDetectionDeclarationProxy
    {
        public EnumMemberConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy)
            : base(identifier, declarationType, accessibility, parentProxy.Prototype)
        {
            ParentProxy = parentProxy;
            ProjectId = parentProxy.ProjectId;
            TargetModule = parentProxy.TargetModule;
        }
    }

    public class ModuleConflictDetectionDeclarationProxy : ConflictDetectionDeclarationProxy, IConflictDetectionModuleDeclarationProxy
    {
        public ModuleConflictDetectionDeclarationProxy(ModuleDeclaration module)
            : this(module.ParentDeclaration, module.QualifiedModuleName.ComponentType, module.IdentifierName)
        {
            _declaration = module;
            _hashCode = module.GetHashCode();
        }

        public ModuleConflictDetectionDeclarationProxy(Declaration project, ComponentType componentType, string identifier)
            :base(identifier, DeclarationType.Module, Accessibility.Public, null)
        {
            _projectID = project.ProjectId;
            ComponentType = componentType;
            TargetModule = null;
            IdentifierName = identifier;
            DeclarationType = DeclarationType.Module;
            Accessibility = Accessibility.Public;
            ParentDeclaration = project;
            _targetModuleName = null;

            _proxyID = Guid.NewGuid().ToString();
            _hashCode = _proxyID.GetHashCode();
        }
    }

    /// <summary>
    /// ConflictDeclarationProxy is a wrapper class for an existing or new <c>Declaration</c> and
    /// is designed for the evaluation of IdentifierName conflicts - not generating a <c>Declaration</c>.
    /// The Proxy class supports manipulating <c>Declaration</c> attributes that would be otherwise readonly.
    /// The manipulated attributes support conflict analysis for proposed renames, relocations, or code insertions.
    /// </summary>
    public class ConflictDetectionDeclarationProxy : IConflictDetectionDeclarationProxy
    {
        protected Declaration _declaration;
        protected int _hashCode;
        protected string _proxyID;

        public ConflictDetectionDeclarationProxy(Declaration prototype, ModuleDeclaration module)
            : this(prototype.IdentifierName, prototype.DeclarationType, prototype.Accessibility, prototype.ParentDeclaration)
        {
            _declaration = prototype;
            TargetModule = module;
            _hashCode = _declaration.GetHashCode();
        }

        public ConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, Declaration parentDeclaration)
        {
            _projectID = parentDeclaration?.ProjectId;
            TargetModule = parentDeclaration?.DeclarationType.HasFlag(DeclarationType.Module) ?? false ? parentDeclaration as ModuleDeclaration : null;
            IdentifierName = identifier;
            DeclarationType = declarationType;
            Accessibility = accessibility;
            ParentDeclaration = parentDeclaration;
            _targetModuleName = TargetModule?.IdentifierName ?? parentDeclaration?.QualifiedModuleName.ComponentName;

            _proxyID = Guid.NewGuid().ToString();

            var hashString = $"{IdentifierName}.{TargetModuleName}.{DeclarationType}.{Accessibility}";
            _hashCode = hashString.GetHashCode();
        }

        public string ProxyID => _proxyID;

        public Declaration Prototype => _declaration;

        public ModuleDeclaration TargetModule { set; get; }

        public QualifiedModuleName? QualifiedModuleName => ParentDeclaration?.QualifiedModuleName;

        public ComponentType ComponentType { set; get; }

        public Declaration ParentDeclaration { set; get; }

        public IConflictDetectionDeclarationProxy ParentProxy { set; get; }

        public bool HasStandardModuleParent => ParentProxy?.ComponentType.Equals(ComponentType.StandardModule) ??
                     Prototype?.QualifiedModuleName.ComponentType.Equals(ComponentType.StandardModule) ??
                     true;

        public string IdentifierName { set; get; }

        public DeclarationType DeclarationType { set; get; }

        public IEnumerable<IdentifierReference> References
            => _declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        protected string _projectID;
        public string ProjectId
        {
            set => _projectID = value;
            get
            {
                return _projectID ?? _declaration?.ProjectId ?? TargetModule?.ProjectId ?? ParentProxy?.ProjectId ?? string.Empty;
            }
        }

        public Accessibility Accessibility { set; get; }

        protected string _targetModuleName;
        public string TargetModuleName
        {
            set => _targetModuleName = value;
            get
            {
                _targetModuleName = TargetModule?.IdentifierName ?? _targetModuleName;
                return _targetModuleName;
            }
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is IConflictDetectionDeclarationProxy proxy))
            {
                return false;
            }

            if (Prototype != null)
            {
                return Prototype == proxy.Prototype;
            }

            if (proxy.ProxyID == ProxyID)
            {
                return true;
            }

            return proxy.ParentDeclaration == ParentDeclaration
                && proxy.ParentProxy?.ProxyID == proxy.ParentProxy?.ProxyID
                && proxy.TargetModuleName == TargetModuleName
                && proxy.DeclarationType == DeclarationType
                && proxy.Accessibility == Accessibility
                && proxy.IdentifierName == IdentifierName;
        }
    }
}
