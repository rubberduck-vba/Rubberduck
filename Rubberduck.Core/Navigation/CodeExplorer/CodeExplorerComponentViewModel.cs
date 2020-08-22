using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerComponentViewModel : CodeExplorerItemViewModel
    {
        public static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Constant, 
            DeclarationType.Enumeration, 
            DeclarationType.Event, 
            DeclarationType.Function, 
            DeclarationType.LibraryFunction, 
            DeclarationType.LibraryProcedure, 
            DeclarationType.Procedure,
            DeclarationType.PropertyGet, 
            DeclarationType.PropertyLet, 
            DeclarationType.PropertySet, 
            DeclarationType.UserDefinedType, 
            DeclarationType.Variable
        };

        private readonly IVBE _vbe;

        public CodeExplorerComponentViewModel(ICodeExplorerNode parent, Declaration declaration, ref List<Declaration> declarations, IVBE vbe) 
            : base(parent, declaration)
        {
            _vbe = vbe;
            SetName();    
            AddNewChildren(ref declarations);
        }

        private string _name;
        public override string Name => _name;

        public override string NameWithSignature => $"{Name}{(IsPredeclared ? " (Predeclared)" : string.Empty)}";

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.ComponentType;

        public bool IsPredeclared => Declaration != null &&
                                     Declaration.IsUserDefined &&
                                     Declaration.DeclarationType == DeclarationType.ClassModule &&
                                     Declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.Document &&
                                     Declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.UserForm &&
                                     Declaration.Attributes.HasPredeclaredIdAttribute(out _);

        public bool IsTestModule => Declaration.DeclarationType == DeclarationType.ProceduralModule
                                    && Declaration.Annotations.Any(annotation => annotation.Annotation is TestModuleAnnotation);

        public override void Synchronize(ref List<Declaration> updated)
        {
            base.Synchronize(ref updated);
            if (Declaration is null)
            {
                return;
            }

            // Document modules might have had the underlying COM object renamed since the last reparse. Let's check...
            SetName();
        }

        protected override void AddNewChildren(ref List<Declaration> updated)
        {
            if (updated is null)
            {
                return;
            }

            var children = updated.Where(declaration =>
                !ReferenceEquals(Declaration, declaration) &&
                declaration.QualifiedModuleName.Equals(Declaration?.QualifiedModuleName)).ToList();

            updated = updated.Except(children.Concat(new [] { Declaration })).ToList();

            foreach (var member in children.Where(declaration => MemberTypes.Contains(declaration.DeclarationType)).ToList())
            {
                AddChild(new CodeExplorerMemberViewModel(this, member, ref children));
            }
        }

        private void SetName()
        {
            _name = Declaration?.IdentifierName ?? string.Empty;

            if (Declaration is null)
            {
                return;
            }

            var qualifiedModuleName = Declaration.QualifiedName.QualifiedModuleName;

            try
            {
                switch (qualifiedModuleName.ComponentType)
                {
                    case ComponentType.Document:
                        using (var app = _vbe?.HostApplication())
                        {
                            var document = app?.GetDocument(qualifiedModuleName);
                            if (document != null)
                            {
                                var parenthesized = document.DocumentName ?? string.Empty;
                                _name = string.IsNullOrEmpty(parenthesized) ? _name : $"{_name} ({parenthesized})";
                            }
                        }

                        break;
                    case ComponentType.ResFile:
                        _name = string.IsNullOrEmpty(_name)
                            ? CodeExplorerUI.CodeExplorer_ResourceFileText
                            : $"{CodeExplorerUI.CodeExplorer_ResourceFileText} ({Path.GetFileName(_name)})";
                        break;
                    case ComponentType.RelatedDocument:
                        _name = string.IsNullOrEmpty(_name) ? string.Empty : Path.GetFileName(_name);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Trace(ex);
            }

            OnNameChanged();
        }
    }
}
