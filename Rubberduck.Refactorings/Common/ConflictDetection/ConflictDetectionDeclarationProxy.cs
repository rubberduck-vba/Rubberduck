using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionDeclarationProxy
    {
        string IdentifierName { set; get; }
        string TargetModuleName { get; }
        ModuleDeclaration TargetModule { set; get; }
        Declaration Prototype { get; }
        DeclarationType DeclarationType { set; get; }
        IEnumerable<IdentifierReference> References { get; }
        Declaration ParentDeclaration { set; get; }
        Accessibility Accessibility { set; get; }
        string ProjectId { get; }
        string ProjectName { get; }
        QualifiedModuleName? QualifiedModuleName { get; }
    }

    public class ConflictDetectionDeclarationProxy : IConflictDetectionDeclarationProxy
    {
        private readonly Declaration _declaration;
        private int _hashCode;

        public ConflictDetectionDeclarationProxy(Declaration prototype, ModuleDeclaration targetModule)
            : this(prototype.IdentifierName, prototype.DeclarationType, prototype.Accessibility, targetModule, targetModule)
        {
            _declaration = prototype;
            ParentDeclaration = _declaration.ParentDeclaration is ModuleDeclaration
                                        ? TargetModule
                                        : _declaration.ParentDeclaration;
        }

        public ConflictDetectionDeclarationProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration)
        {
            TargetModule = targetModule;
            IdentifierName = identifier;
            DeclarationType = declarationType;
            Accessibility = accessibility;
            ParentDeclaration = parentDeclaration;
            _targetModuleName = targetModule?.IdentifierName ?? string.Empty;
            QualifiedModuleName = targetModule?.QualifiedModuleName;

            var test = DeclarationType.ToString();
            var uniqueID = $"{ProjectId}.{_targetModuleName}.{IdentifierName}.{DeclarationType}.{Accessibility}";
            _hashCode = uniqueID.GetHashCode();
            keyID = _hashCode;
        }

        public Declaration Prototype => _declaration;

        public ModuleDeclaration TargetModule { set; get; }

        public QualifiedModuleName? QualifiedModuleName { get; }

        public Declaration ParentDeclaration { set; get; }

        public string IdentifierName { set; get; }

        public int keyID { set; get; }

        public DeclarationType DeclarationType { set; get; }

        public IEnumerable<IdentifierReference> References 
            => _declaration?.References ?? Enumerable.Empty<IdentifierReference>();

        public string ProjectId => _declaration?.ProjectId ?? TargetModule?.ProjectId ?? string.Empty;

        public string ProjectName => _declaration?.ProjectName ?? TargetModule?.ProjectName ?? string.Empty;

        public Accessibility Accessibility { set; get; }

        private string _targetModuleName;
        public string TargetModuleName
        {
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

            return IdentifierName == proxy.IdentifierName
                && DeclarationType == proxy.DeclarationType
                && TargetModuleName == proxy.TargetModuleName
                && ProjectId == proxy.ProjectId
                && Accessibility == proxy.Accessibility;
        }
    }
}
