using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public class DeclarationsFromComProjectLoader : IDeclarationsFromComProjectLoader
    {
        public IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null)
        {
            var declarations = new List<Declaration>();

            var projectName = new QualifiedModuleName(type.Name, type.Path, type.Name, projectId);
            var project = new ProjectDeclaration(type, projectName);
            declarations.Add(project);

            foreach (var alias in type.Aliases.Select(item => new AliasDeclaration(item, project, projectName)))
            {
                declarations.Add(alias);
            }

            foreach (var module in type.Members)
            {
                var moduleIdentifier = module.Type == DeclarationType.Enumeration || module.Type == DeclarationType.UserDefinedType
                    ? $"_{module.Name}"
                    : module.Name;
                var moduleName = new QualifiedModuleName(type.Name, type.Path, moduleIdentifier);

                var moduleDeclarations = GetDeclarationsForModule(module, moduleName, project);
                declarations.AddRange(moduleDeclarations);
            }

            return declarations;
        }

        private static ICollection<Declaration> GetDeclarationsForModule(IComType module, QualifiedModuleName moduleName, ProjectDeclaration project)
        {
            var declarations = new List<Declaration>();

            var attributes = GetModuleAttributes(module);
            var moduleDeclaration = CreateModuleDeclaration(module, moduleName, project, attributes);
            declarations.Add(moduleDeclaration);

            switch (module)
            {
                case IComTypeWithMembers membered:
                    var (memberDeclarations, defaultMember) =
                        GetDeclarationsForProperties(membered.Properties, moduleName, moduleDeclaration);
                    declarations.AddRange(memberDeclarations);
                    AssignDefaultMember(moduleDeclaration, defaultMember);

                    (memberDeclarations, defaultMember) = GetDeclarationsForMembers(
                        membered.Members, 
                        moduleName,
                        moduleDeclaration, 
                        membered.DefaultMember);
                    declarations.AddRange(memberDeclarations);
                    AssignDefaultMember(moduleDeclaration, defaultMember);

                    if (membered is ComCoClass coClass)
                    {
                        (memberDeclarations, defaultMember) = GetDeclarationsForMembers(
                            coClass.SourceMembers,
                            moduleName, 
                            moduleDeclaration, 
                            coClass.DefaultMember, 
                            eventHandlers: true);
                        declarations.AddRange(memberDeclarations);
                        AssignDefaultMember(moduleDeclaration, defaultMember);
                    }

                    break;
                case ComEnumeration enumeration:
                    var enumDeclaration = new Declaration(enumeration, moduleDeclaration, moduleName);
                    declarations.Add(enumDeclaration);
                    var enumerationMembers = enumeration.Members
                        .Select(enumMember => new ValuedDeclaration(enumMember, enumDeclaration, moduleName))
                        .ToList();
                    declarations.AddRange(enumerationMembers);
                    break;
                case ComStruct structure:
                    var typeDeclaration = new Declaration(structure, moduleDeclaration, moduleName);
                    declarations.Add(typeDeclaration);
                    var structMembers = structure.Fields
                        .Select(f => new Declaration(f, typeDeclaration, moduleName))
                        .ToList();
                    declarations.AddRange(structMembers);
                    break;
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
            }

            return declarations;
        }

        private static void AssignDefaultMember(Declaration moduleDeclaration, Declaration defaultMember)
        {
            if (defaultMember != null && moduleDeclaration is ClassModuleDeclaration classDeclaration)
            {
                classDeclaration.DefaultMember = defaultMember;
            }
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
                attributes.AddExtensibleClassAttribute();
            }
            return attributes;
        }

        private static (ICollection<Declaration> memberDeclarations, Declaration defaultMemberDeclaration) GetDeclarationsForMembers(IEnumerable<ComMember> members, QualifiedModuleName moduleName, Declaration moduleDeclaration,
            ComMember defaultMember, bool eventHandlers = false)
        {
            var memberDeclarations = new List<Declaration>();
            Declaration defaultMemberDeclaration = null;

            foreach (var item in members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name)))
            {
                var (memberDeclaration, parameterDeclarations) = GetDeclarationsForMember(moduleName, moduleDeclaration, eventHandlers, item);
                memberDeclarations.Add(memberDeclaration);
                memberDeclarations.AddRange(parameterDeclarations);

                if (moduleDeclaration is ClassModuleDeclaration && item == defaultMember)
                {
                    defaultMemberDeclaration = memberDeclaration;
                }
            }

            return (memberDeclarations, defaultMemberDeclaration);
        }

        private static (Declaration memberDeclaration, List<Declaration> parameterDeclarations) GetDeclarationsForMember(QualifiedModuleName moduleName,
            Declaration parentDeclaration, bool eventHandlers, ComMember item)
        {
            var memberDeclaration = CreateMemberDeclaration(item, moduleName, parentDeclaration, eventHandlers);

            var parameterDeclarations = new List<Declaration>();
            if (memberDeclaration is IParameterizedDeclaration hasParams)
            {
                parameterDeclarations.AddRange(hasParams.Parameters);
            }

            return (memberDeclaration, parameterDeclarations);
        }

        private static (ICollection<Declaration> propertyDeclarations, Declaration propertyMemberDeclaration) GetDeclarationsForProperties(
            IEnumerable<ComField> properties, QualifiedModuleName moduleName, Declaration moduleDeclaration)
        {
            var propertyDeclarations = new List<Declaration>();
            Declaration defaultMemberDeclaration = null;

            foreach (var item in properties.Where(x => !x.Flags.HasFlag(VARFLAGS.VARFLAG_FRESTRICTED)))
            {
                Debug.Assert(item.Type == DeclarationType.Property);
                var attributes = GetPropertyAttibutes(item);
                var (getter, writer) = GetDeclarationsForProperty(moduleName, moduleDeclaration, item, attributes);

                propertyDeclarations.Add(getter);
                if (writer != null)
                {
                    propertyDeclarations.Add(writer);
                }

                if (moduleDeclaration is ClassModuleDeclaration && attributes.HasDefaultMemberAttribute())
                {
                    defaultMemberDeclaration = getter;
                }
            }

            return (propertyDeclarations, defaultMemberDeclaration);
        }

        private static (Declaration getter, Declaration writer) GetDeclarationsForProperty(QualifiedModuleName moduleName, Declaration moduleDeclaration, ComField item, Attributes attributes)
        {
            var getter = new PropertyGetDeclaration(item, moduleDeclaration, moduleName, attributes);

            if (item.Flags.HasFlag(VARFLAGS.VARFLAG_FREADONLY))
            {
                return (getter, null);
            }

            if (item.IsReferenceType)
            {
                var setter = new PropertySetDeclaration(item, moduleDeclaration, moduleName, attributes);
                return (getter, setter);
            }

            var letter = new PropertyLetDeclaration(item, moduleDeclaration, moduleName, attributes);
            return (getter, letter);
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
                    // ReSharper disable once LocalizableElement
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