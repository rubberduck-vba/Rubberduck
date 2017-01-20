//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.IO;
//using System.Linq;
//using System.Reflection;
//using System.Runtime.InteropServices;
//using Rubberduck.Parsing.Annotations;
//using Rubberduck.Parsing.ComReflection;
//using Rubberduck.Parsing.VBA;
//using Rubberduck.VBEditor;

//namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
//{
//    public class RubberduckApiDeclarations : ICustomDeclarationLoader
//    {
//        private readonly RubberduckParserState _state;
//        private readonly List<Declaration> _declarations = new List<Declaration>();

//        public RubberduckApiDeclarations(RubberduckParserState state)
//        {
//            _state = state;
//        }

//        public IReadOnlyList<Declaration> Load()
//        {
//            var assembly = AppDomain.CurrentDomain.GetAssemblies().SingleOrDefault(a => a.GetName().Name.Equals("Rubberduck"));
//            if (assembly == null)
//            {
//                return _declarations;
//            }
            
//            var name = assembly.GetName();
//            var path = Path.ChangeExtension(assembly.Location, "tlb");
            
//            var projectName = new QualifiedModuleName("Rubberduck", path, "Rubberduck");
//            var project = new ProjectDeclaration(projectName.QualifyMemberName("Rubberduck"), "Rubberduck", true, null)
//            {
//                MajorVersion = name.Version.Major, 
//                MinorVersion = name.Version.Minor
//            };
//            _declarations.Add(project);
            
//            var types = assembly.DefinedTypes.WhereIsComVisible();

//            foreach (var type in types)
//            {
//                var module = type.ToModuleDeclaration(project, projectName, type.IsEnum);
//                _declarations.Add(module);

//                var properties = type.GetProperties().WhereIsComVisible().ToList();

//                foreach (var property in properties)
//                {
//                    if (property.CanWrite && property.GetSetMethod().IsPublic)
//                    {
//                        var declaration = property.ToMemberDeclaration(module, false);
//                        _declarations.Add(declaration);
//                    }
//                    if (property.CanRead && property.GetGetMethod().IsPublic)
//                    {
//                        var declaration = property.ToMemberDeclaration(module, true);
//                        _declarations.Add(declaration);
//                    }
//                }

//                var members = type.GetMembers().WhereIsComVisible();

//                foreach (var member in members)
//                {
//                    if (member.MemberType == MemberTypes.Property)
//                    {

//                    }
//                    //var declaration = member.ToMemberDeclaration(project);
//                }
//            };
//            return _declarations;
//        }
//    }

//    internal static class RubberduckApiDeclarationStatics
//    {
//        public static IEnumerable<T> WhereIsComVisible<T>(this IEnumerable<T> source) where T : MemberInfo
//        {
//            return source.Where(member =>
//            {
//                var attr = member.GetCustomAttributes(typeof(ComVisibleAttribute), true).FirstOrDefault();
//                return attr != null && ((ComVisibleAttribute)attr).Value;
//            });
//        }

//        public static Declaration ToModuleDeclaration(this TypeInfo type, Declaration project, QualifiedModuleName projectName, bool isEnum = false)
//        {
//            return isEnum ? new ProceduralModuleDeclaration(projectName.QualifyMemberName(type.Name), project, type.Name, true, null, null) as Declaration :
//                            new ClassModuleDeclaration(projectName.QualifyMemberName(type.Name), project, type.Name, true, null, null);
//        }

//        public static Declaration ToMemberDeclaration(this PropertyInfo member, Declaration parent, bool getter)
//        {
//            if (getter)
//            {
//                return new PropertyGetDeclaration(parent.QualifiedName.QualifiedModuleName.QualifyMemberName(member.Name), 
//                                       parent, 
//                                       parent,
//                                       member.PropertyType.ToVbaTypeName(), 
//                                       null, 
//                                       string.Empty,
//                                       parent.Accessibility,
//                                       null,
//                                       Selection.Home, 
//                                       member.PropertyType.IsArray, 
//                                       true,
//                                       null,
//                                       new Attributes());
//            }
//            if (member.PropertyType.IsClass)
//            {
//                return new PropertySetDeclaration(parent.QualifiedName.QualifiedModuleName.QualifyMemberName(member.Name),
//                                       parent,
//                                       parent,
//                                       member.PropertyType.ToVbaTypeName(),
//                                       parent.Accessibility,
//                                       null,
//                                       Selection.Home,
//                                       true,
//                                       null,
//                                       new Attributes());                
//            }
//            return new PropertyLetDeclaration(parent.QualifiedName.QualifiedModuleName.QualifyMemberName(member.Name),
//                                   parent,
//                                   parent,
//                                   member.PropertyType.ToVbaTypeName(),
//                                   parent.Accessibility,
//                                   null,
//                                   Selection.Home,
//                                   true,
//                                   null,
//                                   new Attributes());
//        }

//        public static string ToVbaTypeName(this Type type)
//        {
//            var name = string.Empty;
//            if (type.IsClass || type.IsEnum)
//            {
//                name = type.Name;               
//            }
//            switch (type.Name)
//            {
//                case "bool":
//                    name = "Boolean";
//                    break;
//                case "short":
//                    name = "Integer";
//                    break;
//                case "int":
//                    name = "Long";
//                    break;
//            }
//            return name + (type.IsArray ? "()" : string.Empty);
//        }
//    }
//}
