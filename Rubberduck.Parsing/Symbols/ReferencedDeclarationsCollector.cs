using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NLog;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public class ReferencedDeclarationsCollector
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly RubberduckParserState _state;

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

        public ReferencedDeclarationsCollector(RubberduckParserState state)
        {
            _state = state;
        }

        private static readonly HashSet<string> IgnoredInterfaceMembers = new HashSet<string> { "QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke" };

        private bool SerializedVersionExists(string name, int major, int minor)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Declarations");
            if (!Directory.Exists(path))
            {
                return false;
            }
            //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
            var testFile = Path.Combine(path, string.Format("{0}.{1}.{2}", name, major, minor) + ".xml");
            if (File.Exists(testFile))
            {
                return true;
            }
            return false;
        }

        private List<Declaration> LoadSerializedBuiltInReferences(string name, int major, int minor)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Declarations");
            var file = Path.Combine(path, string.Format("{0}.{1}.{2}", name, major, minor) + ".xml");
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(file);
            return deserialized.Unwrap();
        }

        public List<Declaration> GetDeclarationsForReference(IReference reference)
        {
            var output = new List<Declaration>();
            var path = reference.FullPath;

            if (SerializedVersionExists(reference.Name, reference.Major, reference.Minor))
            {
                Logger.Trace(string.Format("Deserializing reference '{0}'.", reference.Name));
                return LoadSerializedBuiltInReferences(reference.Name, reference.Major, reference.Minor);
            }
            Logger.Trace(string.Format("COM reflecting reference '{0}'.", reference.Name));

            ITypeLib typeLibrary;
            // Failure to load might mean that it's a "normal" VBProject that will get parsed by us anyway.
            LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out typeLibrary);
            if (typeLibrary == null)
            {
                return output;
            }

            var type = new ComProject(typeLibrary) { Path = path };

            var projectName = new QualifiedModuleName(type.Name, path, type.Name);
            var project = new ProjectDeclaration(type, projectName);

            output.Add(project);
            foreach (var module in type.Members)
            {
                var moduleName = new QualifiedModuleName(reference.Name, path, module.Name);
                
                var attributes = new Attributes();
                if (module.IsPreDeclared)
                {
                    attributes.AddPredeclaredIdTypeAttribute();
                }
                if (module.IsAppObject)
                {
                    attributes.AddGlobalClassAttribute();
                }
                var declaration = CreateModuleDeclaration(module, moduleName, project, attributes);
                output.Add(declaration);
                var membered = module as IComTypeWithMembers;
                if (membered != null)
                {
                    foreach (var item in membered.Members.Where(m => !m.IsRestricted && !IgnoredInterfaceMembers.Contains(m.Name)))
                    {
                        var memberDeclaration = CreateMemberDeclaration(item, moduleName, declaration);
                        output.Add(memberDeclaration);
                        var hasParams = memberDeclaration as IDeclarationWithParameter;
                        if (hasParams != null)
                        {
                            output.AddRange(hasParams.Parameters);
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
                    var enumDeclaration = new Declaration(enumeration, declaration, projectName);
                    var members =
                        enumeration.Members.Select(
                            e =>
                                new Declaration(e, enumDeclaration,
                                    new QualifiedModuleName(reference.Name, path, enumeration.Name)));
                    output.Add(enumDeclaration);
                    output.AddRange(members);
                }

                var fields = module as IComTypeWithFields;
                if (fields == null || !fields.Fields.Any())
                {
                    continue;
                }
                var declarations = fields.Fields.Select(f => new Declaration(f, declaration, projectName));
                output.AddRange(declarations);
            }
            return output;
        }

        private Declaration CreateModuleDeclaration(IComType module, QualifiedModuleName project, Declaration parent, Attributes attributes)
        {
            var enumeration = module as ComEnumeration;
            if (enumeration != null)
            {
                return new ProceduralModuleDeclaration(enumeration, parent, project);
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
            var normal = module as ComModule;
            if (normal != null && normal.Fields.Any())
            {
                //These are going to be UDTs or Consts.  Apparently COM modules can declare *either* fields or members.
                return new ProceduralModuleDeclaration(string.Format("_{0}", normal.Name), parent, project);
            }
            return new ProceduralModuleDeclaration(normal, parent, project, attributes);
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
                    Debug.Assert(false);
                    return null as Declaration;
            }
        }
    }
}
