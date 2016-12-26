using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
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
        private SerializableProject _serialized;
        private readonly Dictionary<Declaration, SerializableDeclarationTree> _treeLookup = new Dictionary<Declaration, SerializableDeclarationTree>(); 
        private readonly List<Declaration> _declarations = new List<Declaration>(); 

        private const string EnumPseudoName = "Enums";
        private Declaration _enumModule;
        private const string TypePseudoName = "Types";        
        private Declaration _typeModule;

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

        public ReferencedDeclarationsCollector(RubberduckParserState state, IReference reference)
        {
            _state = state;
            _path = reference.FullPath;
            _referenceName = reference.Name;
            _referenceMajor = reference.Major;
            _referenceMinor = reference.Minor;
        }
        
        public bool SerializedVersionExists
        {
            get
            {
                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Declarations");
                if (!Directory.Exists(path))
                {
                    return false;
                }
                //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
                var testFile = Path.Combine(path, string.Format("{0}.{1}.{2}", _referenceName, _referenceMajor, _referenceMinor) + ".xml");
                return File.Exists(testFile);
            }
        }

        public List<Declaration> LoadDeclarationsFromXml()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Declarations");
            var file = Path.Combine(path, string.Format("{0}.{1}.{2}", _referenceName, _referenceMajor, _referenceMinor) + ".xml");
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(file);
            return deserialized.Unwrap();
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

            foreach (var module in type.Members)
            {
                var moduleName = new QualifiedModuleName(_referenceName, _path, module.Name);

                var attributes = new Attributes();
                if (module.IsPreDeclared)
                {
                    attributes.AddPredeclaredIdTypeAttribute();
                }
                if (module.IsAppObject)
                {
                    attributes.AddGlobalClassAttribute();
                }

                var declaration = CreateModuleDeclaration(module,
                    module.Type == DeclarationType.Enumeration || module.Type == DeclarationType.UserDefinedType
                        ? projectName
                        : moduleName, project, attributes);

                if (declaration.IdentifierName.Equals(EnumPseudoName))
                {
                    if (_enumModule == null)
                    {
                        _enumModule = declaration;
                        AddToOutput(_enumModule, null);
                    }
                }
                else if (declaration.IdentifierName.Equals(TypePseudoName))
                {
                    if (_typeModule == null)
                    {
                        _typeModule = declaration;
                        AddToOutput(_typeModule, null);
                    }
                }
                else
                {
                    AddToOutput(declaration, null);
                }   

                var membered = module as IComTypeWithMembers;
                if (membered != null)
                {
                    foreach (var item in membered.Members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name)))
                    {
                        var memberDeclaration = CreateMemberDeclaration(item, moduleName, declaration);
                        AddToOutput(memberDeclaration, declaration);
                        var hasParams = memberDeclaration as IDeclarationWithParameter;
                        if (hasParams != null)
                        {
                            AddRangeToOutput(hasParams.Parameters, memberDeclaration);
                        }
                        var coClass = memberDeclaration as ClassModuleDeclaration;
                        if (coClass != null && item.IsDefault)
                        {
                            coClass.DefaultMember = memberDeclaration;
                        }
                    }
                }

                var enumeration = module as ComEnumeration;
                if (enumeration != null)
                {
                    var qualified = new QualifiedModuleName(_referenceName, _path, EnumPseudoName);
                    var enumDeclaration = new Declaration(enumeration, declaration, qualified);
                    var members = enumeration.Members.Select(e => new Declaration(e, enumDeclaration, qualified));
                    AddToOutput(enumDeclaration, null);
                    AddRangeToOutput(members, enumDeclaration);
                }

                var structure = module as ComStruct;
                if (structure != null)
                {
                    var qualified = new QualifiedModuleName(_referenceName, _path, TypePseudoName);
                    var typeDeclaration = new Declaration(structure, declaration, qualified);
                    var members = structure.Fields.Select(f => new Declaration(f, typeDeclaration, qualified));
                    AddToOutput(typeDeclaration, null);
                    AddRangeToOutput(members, typeDeclaration);
                }

                var fields = module as IComTypeWithFields;
                if (fields == null || !fields.Fields.Any())
                {
                    continue;
                }
                var declarations = fields.Fields.Select(f => new Declaration(f, declaration, projectName));
                AddRangeToOutput(declarations, declaration);
            }
            _state.BuiltInDeclarationTrees.TryAdd(_serialized);
            return _declarations;
        }

        private void AddToOutput(Declaration declaration, Declaration parent)
        {
            _declarations.Add(declaration);
            //if (parent == null)
            //{
            //    var tree = new SerializableDeclarationTree(declaration);
            //    _treeLookup.Add(declaration, tree);
            //    _serialized.AddDeclaration(tree);
            //}
            //else
            //{
            //    var tree = new SerializableDeclarationTree(declaration);
            //    Debug.Assert(!_treeLookup.ContainsKey(declaration));
            //    _treeLookup.Add(declaration, tree);
            //    _treeLookup[parent].AddChildTree(tree);                
            //}
        }

        private void AddRangeToOutput(IEnumerable<Declaration> declarations, Declaration parent)
        {
            //Debug.Assert(parent != null);
            //var tree = _treeLookup[parent];
            foreach (var declaration in declarations)
            {
            //    Debug.Assert(!_treeLookup.ContainsKey(declaration));
            //    var child = new SerializableDeclarationTree(declaration);
            //    _treeLookup.Add(declaration, child);
            //    tree.AddChildTree(child);
                _declarations.Add(declaration);
            }
        }

        private Declaration CreateModuleDeclaration(IComType module, QualifiedModuleName project, Declaration parent, Attributes attributes)
        {
            var enumeration = module as ComEnumeration;
            if (enumeration != null)
            {
                //There's no real reason that these can't all live in one pseudo-module.
                return _enumModule ?? new ProceduralModuleDeclaration(EnumPseudoName, parent, project);
            }
            var types = module as ComStruct;
            if (types != null)
            {
                //There's also no real reason that *these* can't all live in one pseudo-module.
                return _typeModule ?? new ProceduralModuleDeclaration(TypePseudoName, parent, project);
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

        private Declaration CreateMemberDeclaration(ComMember member, QualifiedModuleName module, Declaration parent)
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

            switch (member.Type)
            {
                case DeclarationType.Event:
                case DeclarationType.Procedure:
                    return new SubroutineDeclaration(member, parent, module, attributes);
                case DeclarationType.Function:
                    return new FunctionDeclaration(member, parent, module, attributes);
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

        //private SerializableProject GetSerializableProject(ProjectDeclaration declaration, List<Declaration> declarations)
        //{
        //    var project = new SerializableProject(declaration);
        //    var children = new List<SerializableDeclarationTree>();
        //    var nodes = declarations.Where(x => x.ParentDeclaration.Equals(declaration)).ToList();
        //    foreach (var item in nodes)
        //    {
        //        children.Add(GetSerializableTreeForDeclaration(item, declarations));
        //    }
        //    project.Declarations = children;
        //    return project;
        //}

        //private SerializableDeclarationTree GetSerializableTreeForDeclaration(Declaration declaration, List<Declaration> declarations)
        //{
        //    var children = new List<SerializableDeclarationTree>();
        //    var nodes = declarations.Where(x => x.ParentDeclaration.Equals(declaration)).ToList();
        //    declarations.RemoveAll(nodes.Contains);
        //    foreach (var item in nodes)
        //    {
        //        children.Add(GetSerializableTreeForDeclaration(item, declarations));
        //    }
        //    return new SerializableDeclarationTree(declaration, children);
        //}
    }
}
