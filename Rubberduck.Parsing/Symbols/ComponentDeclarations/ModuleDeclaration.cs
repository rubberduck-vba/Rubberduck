using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public abstract class ModuleDeclaration : Declaration
    {
        protected ModuleDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            DeclarationType declarationType,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes,
            bool isWithEvents = false)
            : base(
                qualifiedName,
                projectDeclaration,
                projectDeclaration,
                name,
                null,
                false,
                isWithEvents,
                Accessibility.Public,
                declarationType,
                null,
                null,
                Selection.Home,
                false,
                null,
                isUserDefined,
                annotations,
                attributes)
        {
            CustomFolder = FolderFromAnnotations();
        }

        private readonly List<Declaration> _members = new List<Declaration>();
        public IReadOnlyList<Declaration> Members => _members;

        internal void AddMember(Declaration member)
        {
            _members.Add(member);
        }

        internal void RemoveAnnotations(ICollection<IAnnotation> annotationsToRemove)
        {
            _annotations = _annotations?.Where(annotation => !annotationsToRemove.Contains(annotation)).ToList();
        }

        public override string CustomFolder { get; }

        private string FolderFromAnnotations()
        {
            var @namespace = Annotations.FirstOrDefault(annotation => annotation.AnnotationType == AnnotationType.Folder);
            string result;
            if (@namespace == null)
            {
                result = string.IsNullOrEmpty(QualifiedName.QualifiedModuleName.ProjectName)
                    ? ProjectId
                    : QualifiedName.QualifiedModuleName.ProjectName;
            }
            else
            {
                var value = ((FolderAnnotation)@namespace).FolderName;
                result = value;
            }
            return result;
        }
    }
}
