using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.ComReflection
{
    public class ReferencedDeclarationsCollector
    {
        #region Native Stuff
        // ReSharper disable InconsistentNaming
        // ReSharper disable UnusedMember.Local
        /// <summary>
        /// Controls how a type library is registered.
        /// </summary>
        private enum REGKIND
        {
            /// <summary>
            /// Use default register behavior.
            /// </summary>


            REGKIND_DEFAULT = 0,
            /// <summary>
            /// Register this type library.
            /// </summary>
            REGKIND_REGISTER = 1,
            /// <summary>
            /// Do not register this type library.
            /// </summary>
            REGKIND_NONE = 2
        }
        // ReSharper restore UnusedMember.Local

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        private static extern int LoadTypeLibEx(string strTypeLibName, REGKIND regKind, out ITypeLib TypeLib);
        // ReSharper restore InconsistentNaming
        #endregion

        private readonly RubberduckParserState _state;
        private readonly string _serializedDeclarationsPath;
        private SerializableProject _serialized;
        private readonly List<Declaration> _declarations = new List<Declaration>(); 

        private static readonly HashSet<string> IgnoredInterfaceMembers = new HashSet<string>
        {
            "QueryInterface",
            "AddRef",
            "Release",
            "GetTypeInfoCount",
            "GetTypeInfo",
            "GetIDsOfNames",
            "Invoke"
        };

        private readonly string _referenceName;
        private readonly string _path;
        private readonly int _referenceMajor;
        private readonly int _referenceMinor;

        public ReferencedDeclarationsCollector(RubberduckParserState state, IReference reference, string serializedDeclarationsPath)
        {
            _state = state;
            _serializedDeclarationsPath = serializedDeclarationsPath;
            _path = reference.FullPath;
            _referenceName = reference.Name;
            _referenceMajor = reference.Major;
            _referenceMinor = reference.Minor;
        }
        
        public bool SerializedVersionExists
        {
            get
            {
                if (!Directory.Exists(_serializedDeclarationsPath))
                {
                    return false;
                }
                //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
                var testFile = Path.Combine(_serializedDeclarationsPath, string.Format("{0}.{1}.{2}", _referenceName, _referenceMajor, _referenceMinor) + ".xml");
                return File.Exists(testFile);
            }
        }

        private static readonly HashSet<DeclarationType> ProceduralTypes =
            new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet,
                DeclarationType.PropertyLet, DeclarationType.PropertySet
            });

        public List<Declaration> LoadDeclarationsFromXml()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Declarations");
            var file = Path.Combine(path, string.Format("{0}.{1}.{2}", _referenceName, _referenceMajor, _referenceMinor) + ".xml");
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(file);

            var declarations = deserialized.Unwrap();

            foreach (var members in declarations.Where(d => d.DeclarationType != DeclarationType.Project && 
                                                            d.ParentDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                            ProceduralTypes.Contains(d.DeclarationType))
                                                .GroupBy(d => d.ParentDeclaration))
            { 
                _state.CoClasses.TryAdd(members.Select(m => m.IdentifierName).ToList(), members.First().ParentDeclaration);
            }
            return declarations;
        }

        public List<Declaration> LoadDeclarationsFromLibrary()
        {
            ITypeLib typeLibrary;
            // Failure to load might mean that it's a "normal" VBProject that will get parsed by us anyway.
            LoadTypeLibEx(_path, REGKIND.REGKIND_NONE, out typeLibrary);
            if (typeLibrary == null)
            {
                return _declarations;
            }

            var type = new ComProject(typeLibrary) { Path = _path };

            var projectName = new QualifiedModuleName(type.Name, _path, type.Name);
            var project = new ProjectDeclaration(type, projectName);
            _serialized = new SerializableProject(project);
            _declarations.Add(project);

            foreach (var alias in type.Aliases.Select(item => new AliasDeclaration(item, project, projectName)))
            {
                _declarations.Add(alias);
                _serialized.AddDeclaration(new SerializableDeclarationTree(alias));
            }

            foreach (var module in type.Members)
            {
                var moduleName = new QualifiedModuleName(_referenceName, _path,
                    module.Type == DeclarationType.Enumeration || module.Type == DeclarationType.UserDefinedType
                        ? string.Format("_{0}", module.Name)
                        : module.Name);

                var declaration = CreateModuleDeclaration(module, moduleName, project, GetModuleAttributes(module));
                var moduleTree = new SerializableDeclarationTree(declaration);
                _declarations.Add(declaration);
                _serialized.AddDeclaration(moduleTree);

                var membered = module as IComTypeWithMembers;
                if (membered != null)
                {
                    CreateMemberDeclarations(membered.Members, moduleName, declaration, moduleTree, membered.DefaultMember);
                    var coClass = membered as ComCoClass;
                    if (coClass != null)
                    {
                        CreateMemberDeclarations(coClass.SourceMembers, moduleName, declaration, moduleTree, coClass.DefaultMember, true);
                    }
                }

                var enumeration = module as ComEnumeration;
                if (enumeration != null)
                {
                    var enumDeclaration = new Declaration(enumeration, declaration, moduleName);
                    _declarations.Add(enumDeclaration);
                    var members = enumeration.Members.Select(e => new Declaration(e, enumDeclaration, moduleName)).ToList();
                    _declarations.AddRange(members);

                    var enumTree = new SerializableDeclarationTree(enumDeclaration);
                    moduleTree.AddChildTree(enumTree);
                    enumTree.AddChildren(members);
                }

                var structure = module as ComStruct;
                if (structure != null)
                {
                    var typeDeclaration = new Declaration(structure, declaration, moduleName);
                    _declarations.Add(typeDeclaration);
                    var members = structure.Fields.Select(f => new Declaration(f, typeDeclaration, moduleName)).ToList();
                    _declarations.AddRange(members);

                    var typeTree = new SerializableDeclarationTree(typeDeclaration);
                    moduleTree.AddChildTree(typeTree);
                    typeTree.AddChildren(members);
                }

                var fields = module as IComTypeWithFields;
                if (fields == null || !fields.Fields.Any())
                {
                    continue;
                }
                var declarations = fields.Fields.Select(f => new Declaration(f, declaration, projectName)).ToList();
                _declarations.AddRange(declarations);
                moduleTree.AddChildren(declarations);
            }
            _state.BuiltInDeclarationTrees.TryAdd(_serialized);
            return _declarations;
        }

        private static Attributes GetModuleAttributes(IComType module)
        {
            var attributes = new Attributes();
            if (module.IsPreDeclared)
            {
                attributes.AddPredeclaredIdTypeAttribute();
            }
            if (module.IsAppObject)
            {
                attributes.AddGlobalClassAttribute();
            }
            if (module as IComTypeWithMembers != null && ((IComTypeWithMembers)module).IsExtensible)
            {
                attributes.AddExtensibledClassAttribute();
            }
            return attributes;
        }

        private void CreateMemberDeclarations(IEnumerable<ComMember> members, QualifiedModuleName moduleName, Declaration declaration,
                                              SerializableDeclarationTree moduleTree, ComMember defaultMember, bool eventHandlers = false)
        {
            foreach (var item in members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name)))
            {
                var memberDeclaration = CreateMemberDeclaration(item, moduleName, declaration, eventHandlers);
                _declarations.Add(memberDeclaration);

                var memberTree = new SerializableDeclarationTree(memberDeclaration);
                moduleTree.AddChildTree(memberTree);

                var hasParams = memberDeclaration as IParameterizedDeclaration;
                if (hasParams != null)
                {
                    _declarations.AddRange(hasParams.Parameters);
                    memberTree.AddChildren(hasParams.Parameters);
                }
                var coClass = memberDeclaration as ClassModuleDeclaration;
                if (coClass != null && item == defaultMember)
                {
                    coClass.DefaultMember = memberDeclaration;
                }
            }
        }

        private Declaration CreateModuleDeclaration(IComType module, QualifiedModuleName project, Declaration parent, Attributes attributes)
        {
            var enumeration = module as ComEnumeration;
            if (enumeration != null)
            {
                return new ProceduralModuleDeclaration(enumeration, parent, project);
            }
            var types = module as ComStruct;
            if (types != null)
            {
                return new ProceduralModuleDeclaration(types, parent, project);
            }
            var coClass = module as ComCoClass;
            var intrface = module as ComInterface;
            if (coClass != null || intrface != null)
            {
                var output = coClass != null ? 
                    new ClassModuleDeclaration(coClass, parent, project, attributes) :
                    new ClassModuleDeclaration(intrface, parent, project, attributes);
                if (coClass != null)
                {
                    var members =
                        coClass.Members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name))
                            .Select(m => m.Name);
                    _state.CoClasses.TryAdd(members.ToList(), output);
                }
                return output;
            }
            return new ProceduralModuleDeclaration(module as ComModule, parent, project, attributes);
        }

        private Declaration CreateMemberDeclaration(ComMember member, QualifiedModuleName module, Declaration parent, bool handler)
        {
            var attributes = GetMemberAttibutes(member);
            switch (member.Type)
            {
                case DeclarationType.Procedure:
                    return new SubroutineDeclaration(member, parent, module, attributes, handler);
                case DeclarationType.Function:
                    return new FunctionDeclaration(member, parent, module, attributes);
                case DeclarationType.Event:
                    return new EventDeclaration(member, parent, module, attributes);
                case DeclarationType.PropertyGet:
                    return new PropertyGetDeclaration(member, parent, module, attributes);
                case DeclarationType.PropertySet:
                    return new PropertySetDeclaration(member, parent, module, attributes);
                case DeclarationType.PropertyLet:
                    return new PropertyLetDeclaration(member, parent, module, attributes);
                default:
                    throw new InvalidEnumArgumentException(string.Format("Unexpected DeclarationType {0} encountered.", member.Type));
            }
        }

        private static Attributes GetMemberAttibutes(ComMember member)
        {
            var attributes = new Attributes();
            if (member.IsEnumerator)
            {
                attributes.AddEnumeratorMemberAttribute(member.Name);
            }
            else if (member.IsDefault)
            {
                attributes.AddDefaultMemberAttribute(member.Name);
            }
            else if (member.IsHidden)
            {
                attributes.AddHiddenMemberAttribute(member.Name);
            }
            else if (member.IsEvaluateFunction)
            {
                attributes.AddEvaluateMemberAttribute(member.Name);
            }
            else if (!string.IsNullOrEmpty(member.Documentation.DocString))
            {
                attributes.AddMemberDescriptionAttribute(member.Name, member.Documentation.DocString);
            }
            return attributes;
        }
    }
}
