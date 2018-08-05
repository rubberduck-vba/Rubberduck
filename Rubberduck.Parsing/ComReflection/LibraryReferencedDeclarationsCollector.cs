using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace Rubberduck.Parsing.ComReflection
{
    public class LibraryReferencedDeclarationsCollector : IReferencedDeclarationsCollector
    {
        private readonly IComLibraryProvider _comLibraryProvider;

        public LibraryReferencedDeclarationsCollector(IComLibraryProvider comLibraryProvider)
        {
            _comLibraryProvider = comLibraryProvider;
        }

        public (IReadOnlyCollection<Declaration> declarations, Dictionary<IList<string>, Declaration> coClasses, SerializableProject serializableProject) CollectDeclarations(
            IReference reference)
        {
            return LoadDeclarationsFromLibrary(reference);
        }

        private (List<Declaration> declarations, Dictionary<IList<string>, Declaration> coClasses, SerializableProject serializableProject) LoadDeclarationsFromLibrary(IReference reference)
        {
            var libraryPath = reference.FullPath;
            // Failure to load might mean that it's a "normal" VBProject that will get parsed by us anyway.
            var typeLibrary = GetTypeLibrary(libraryPath);
            if (typeLibrary == null)
            {
                return (new List<Declaration>(), new Dictionary<IList<string>, Declaration>(), null) ;
            }

            var declarations = new List<Declaration>();
            var coClasses = new Dictionary<IList<string>,Declaration>();

            var type = new ComProject(typeLibrary, libraryPath);

            var projectName = new QualifiedModuleName(type.Name, libraryPath, type.Name);
            var project = new ProjectDeclaration(type, projectName);
            var serialized = new SerializableProject(project);
            declarations.Add(project);

            foreach (var alias in type.Aliases.Select(item => new AliasDeclaration(item, project, projectName)))
            {
                declarations.Add(alias);
                serialized.AddDeclaration(new SerializableDeclarationTree(alias));
            }

            foreach (var module in type.Members)
            {
                var moduleIdentifier = module.Type == DeclarationType.Enumeration || module.Type == DeclarationType.UserDefinedType
                                        ? $"_{module.Name}"
                                        : module.Name;
                var moduleName = new QualifiedModuleName(reference.Name, libraryPath, moduleIdentifier);

                var (moduleDeclarations, coClass, moduleTree) = GetDeclarationsForModule(module, moduleName, project);
                declarations.AddRange(moduleDeclarations);
                if (coClass.HasValue)
                {
                    coClasses[coClass.Value.Key] = coClass.Value.Value;
                }
                serialized.AddDeclaration(moduleTree);
            }
            return (declarations, coClasses, serialized);
        }

        private static (ICollection<Declaration> declarations, KeyValuePair<IList<string>, Declaration>? coClass, SerializableDeclarationTree moduleTree) GetDeclarationsForModule(IComType module, QualifiedModuleName moduleName,
            ProjectDeclaration project)
        {
            var declarations = new List<Declaration>();
            KeyValuePair<IList<string>, Declaration>? coClassItem = null;

            var attributes = GetModuleAttributes(module);
            var moduleDeclaration = CreateModuleDeclaration(module, moduleName, project, attributes);
            var moduleTree = new SerializableDeclarationTree(moduleDeclaration);
            declarations.Add(moduleDeclaration);

            switch (module)
            {
                case IComTypeWithMembers membered:
                    var (memberDeclarations, defaultMember, memberTrees) =
                        GetDeclarationsForProperties(membered.Properties, moduleName, moduleDeclaration);
                    declarations.AddRange(memberDeclarations);
                    moduleTree.AddChildTrees(memberTrees);
                    AssignDefaultMember(moduleDeclaration, defaultMember);

                    (memberDeclarations, defaultMember, memberTrees) = GetDeclarationsForMembers(membered.Members, moduleName,
                        moduleDeclaration, membered.DefaultMember);
                    declarations.AddRange(memberDeclarations);
                    moduleTree.AddChildTrees(memberTrees);
                    AssignDefaultMember(moduleDeclaration, defaultMember);

                    if (membered is ComCoClass coClass)
                    {
                        var memberList = CoClassMemberList(coClass);
                        coClassItem = new KeyValuePair<IList<string>, Declaration>(memberList, moduleDeclaration);

                        (memberDeclarations, defaultMember, memberTrees) = GetDeclarationsForMembers(coClass.SourceMembers,
                            moduleName, moduleDeclaration, coClass.DefaultMember, true);
                        declarations.AddRange(memberDeclarations);
                        moduleTree.AddChildTrees(memberTrees);
                        AssignDefaultMember(moduleDeclaration, defaultMember);
                    }

                    break;
                case ComEnumeration enumeration:
                {
                    var enumDeclaration = new Declaration(enumeration, moduleDeclaration, moduleName);
                    declarations.Add(enumDeclaration);
                    var members = enumeration.Members.Select(e => new ValuedDeclaration(e, enumDeclaration, moduleName))
                        .ToList();
                    declarations.AddRange(members);

                    var enumTree = new SerializableDeclarationTree(enumDeclaration);
                    moduleTree.AddChildTree(enumTree);
                    enumTree.AddChildren(members);
                    break;
                }
                case ComStruct structure:
                {
                    var typeDeclaration = new Declaration(structure, moduleDeclaration, moduleName);
                    declarations.Add(typeDeclaration);
                    var members = structure.Fields.Select(f => new Declaration(f, typeDeclaration, moduleName)).ToList();
                    declarations.AddRange(members);

                    var typeTree = new SerializableDeclarationTree(typeDeclaration);
                    moduleTree.AddChildTree(typeTree);
                    typeTree.AddChildren(members);
                    break;
                }
            }

            if (module is IComTypeWithFields fields && fields.Fields.Any())
            {
                var projectName = project.QualifiedModuleName;
                var fieldDeclarations = new List<Declaration>();
                foreach (var field in fields.Fields)
                {
                    fieldDeclarations.Add(field.Type == DeclarationType.Constant
                        ? new ValuedDeclaration(field, moduleDeclaration, projectName)
                        : new Declaration(field, moduleDeclaration, projectName));
                }

                declarations.AddRange(fieldDeclarations);
                moduleTree.AddChildren(fieldDeclarations);
            }

            return (declarations, coClassItem, moduleTree);
        }

        private static void AssignDefaultMember(Declaration moduleDeclaration, Declaration defaultMember)
        {
            if (defaultMember != null && moduleDeclaration is ClassModuleDeclaration classDeclaration)
            {
                classDeclaration.DefaultMember = defaultMember;
            }
        }

        private static List<string> CoClassMemberList(ComCoClass comCoClass)
        {
            return comCoClass.Members
                .Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name))
                .Select(m => m.Name)
                .ToList();
        }

        private ITypeLib GetTypeLibrary(string libraryPath)
        {
            return _comLibraryProvider.LoadTypeLibrary(libraryPath);
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
            if (module is IComTypeWithMembers members && members.IsExtensible)
            {
                attributes.AddExtensibledClassAttribute();
            }
            return attributes;
        }

        private static (ICollection<Declaration> memberDeclarations, Declaration defaultMemberDeclaration, ICollection<SerializableDeclarationTree> memberTrees) GetDeclarationsForMembers(IEnumerable<ComMember> members, QualifiedModuleName moduleName, Declaration moduleDeclaration,
            ComMember defaultMember, bool eventHandlers = false)
        {
            var memberDeclarations = new List<Declaration>();
            var memberTrees = new List<SerializableDeclarationTree>();
            Declaration defaultMemberDeclaration = null;

            foreach (var item in members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name)))
            {
                var (memberDeclaration, parameterDeclarations, memberTree) = GetDeclarationsForMember(moduleName, moduleDeclaration, eventHandlers, item);
                memberDeclarations.Add(memberDeclaration);
                memberDeclarations.AddRange(parameterDeclarations);

                if (moduleDeclaration is ClassModuleDeclaration && item == defaultMember)
                {
                    defaultMemberDeclaration = memberDeclaration;
                }

                memberTrees.Add(memberTree);
            }

            return (memberDeclarations, defaultMemberDeclaration, memberTrees);
        }

        private static (Declaration memberDeclaration, ICollection<Declaration> parameterDeclarations, SerializableDeclarationTree memberTree) GetDeclarationsForMember(QualifiedModuleName moduleName, 
            Declaration declaration, bool eventHandlers, ComMember item)
        {
            var memberDeclaration = CreateMemberDeclaration(item, moduleName, declaration, eventHandlers);
            var memberTree = new SerializableDeclarationTree(memberDeclaration);

            var parameterDeclarations = new List<Declaration>();
            if (memberDeclaration is IParameterizedDeclaration hasParams)
            {
                parameterDeclarations.AddRange(hasParams.Parameters);
                memberTree.AddChildren(hasParams.Parameters);
            }

            return (memberDeclaration, parameterDeclarations, memberTree);
        }

        private static (ICollection<Declaration> propertyDeclarations, Declaration propertyMemberDeclaration, ICollection<SerializableDeclarationTree> propertyTrees) GetDeclarationsForProperties(
            IEnumerable<ComField> properties, QualifiedModuleName moduleName, Declaration moduleDeclaration)
        {
            var propertyDeclarations = new List<Declaration>();
            var propertyTrees = new List<SerializableDeclarationTree>();
            Declaration defaultMemberDeclaration = null;

            foreach (var item in properties.Where(x => !x.Flags.HasFlag(VARFLAGS.VARFLAG_FRESTRICTED)))
            {
                Debug.Assert(item.Type == DeclarationType.Property);
                var attributes = GetPropertyAttibutes(item);
                var (getter, writer, itemPropertyTrees) = GetDeclarationsForProperty(moduleName, moduleDeclaration, item, attributes);

                propertyDeclarations.Add(getter);
                if (writer != null)
                {
                    propertyDeclarations.Add(writer);
                }

                if (moduleDeclaration is ClassModuleDeclaration && attributes.HasDefaultMemberAttribute())
                {
                    defaultMemberDeclaration = getter;
                }

                propertyTrees.AddRange(itemPropertyTrees);
            }

            return (propertyDeclarations, defaultMemberDeclaration, propertyTrees);
        }

        private static (Declaration getter, Declaration writer, ICollection<SerializableDeclarationTree> propertyTrees) GetDeclarationsForProperty(QualifiedModuleName moduleName, Declaration moduleDeclaration, ComField item, Attributes attributes)
        {
            var propertyTrees = new List<SerializableDeclarationTree>();

            var getter = new PropertyGetDeclaration(item, moduleDeclaration, moduleName, attributes);
            var getterTree = new SerializableDeclarationTree(getter);
            propertyTrees.Add(getterTree);

            if (item.Flags.HasFlag(VARFLAGS.VARFLAG_FREADONLY))
            {
                return (getter, null, propertyTrees);
            }

            if (item.IsReferenceType)
            {
                var setter = new PropertySetDeclaration(item, moduleDeclaration, moduleName, attributes);
                var setterTree = new SerializableDeclarationTree(setter);
                propertyTrees.Add(setterTree);
                return (getter, setter, propertyTrees);
            }

            var letter = new PropertyLetDeclaration(item, moduleDeclaration, moduleName, attributes);
            var letterTree = new SerializableDeclarationTree(letter);
            propertyTrees.Add(letterTree);
            return (getter, letter, propertyTrees);
        }

        private static Declaration CreateModuleDeclaration(IComType module, QualifiedModuleName project, Declaration parent, Attributes attributes)
        {
            switch (module)
            {
                case ComEnumeration enumeration:
                    return new ProceduralModuleDeclaration(enumeration, parent, project);
                case ComStruct types:
                    return new ProceduralModuleDeclaration(types, parent, project);
                case ComCoClass coClass:
                    return new ClassModuleDeclaration(coClass, parent, project, attributes);
                case ComInterface intrface:
                    return new ClassModuleDeclaration(intrface, parent, project, attributes);
                default:
                    return new ProceduralModuleDeclaration(module as ComModule, parent, project, attributes);
            }
        }

        private static Declaration CreateMemberDeclaration(ComMember member, QualifiedModuleName module, Declaration parent, bool handler)
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
                    throw new InvalidEnumArgumentException($"Unexpected DeclarationType {member.Type} encountered.");
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

        private static Attributes GetPropertyAttibutes(ComField property)
        {
            var attributes = new Attributes();
            if (property.Flags.HasFlag(VARFLAGS.VARFLAG_FDEFAULTBIND))
            {
                attributes.AddDefaultMemberAttribute(property.Name);
            }
            if (property.Flags.HasFlag(VARFLAGS.VARFLAG_FHIDDEN))
            {
                attributes.AddHiddenMemberAttribute(property.Name);
            }
            return attributes;
        }

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
    }
}
