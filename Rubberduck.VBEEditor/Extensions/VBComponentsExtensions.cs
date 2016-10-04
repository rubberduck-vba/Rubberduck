using System;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

// todo: untangle this mess

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBComponentsExtensions
    {
        /// <summary>
        /// Safely removes the specified VbComponent from the collection.
        /// </summary>
        /// <remarks>
        /// UserForms, Class modules, and Standard modules are completely removed from the project.
        /// Since Document type components can't be removed through the VBE, all code in its CodeModule are deleted instead.
        /// </remarks>
        public static void RemoveSafely(this VBComponents components, VBComponent component)
        {
            switch (component.Type)
            {
                case ComponentType.ClassModule:
                case ComponentType.StandardModule:
                case ComponentType.UserForm:
                    components.Remove(component);
                    break;
                case ComponentType.ActiveXDesigner:
                case ComponentType.Document:
                    component.CodeModule.Clear();
                    break;
                default:
                    break;
            }
        }

        public static void ImportSourceFile(this VBComponents components, string filePath)
        {
            var ext = Path.GetExtension(filePath);
            var name = Path.GetFileNameWithoutExtension(filePath);
            if (!File.Exists(filePath))
            {
                return;
            }

            var codeString = File.ReadAllText(filePath);
            var codeLines = codeString.Split(new []{Environment.NewLine}, StringSplitOptions.None);
            if (ext == VBComponentExtensions.DocClassExtension)
            {
                var component = components.Item(name);
                component.CodeModule.Clear();
                component.CodeModule.AddFromString(codeString);
            }
            else if (ext == VBComponentExtensions.FormExtension)
            {
                VBComponent component;
                try
                {
                    component = components.Item(name);
                }
                catch (IndexOutOfRangeException)
                {
                    component = components.Add(ComponentType.UserForm);
                    component.Properties.Item("Caption").Value = name;
                    component.Name = name;
                }

                var nonAttributeLines = codeLines.TakeWhile(line => !line.StartsWith("Attribute")).Count();
                var attributeLines = codeLines.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
                var declarationsStartLine = nonAttributeLines + attributeLines + 1;
                var correctCodeString = string.Join(Environment.NewLine, codeLines.Skip(declarationsStartLine - 1).ToArray());

                component.CodeModule.Clear();
                component.CodeModule.AddFromString(correctCodeString);
            }
            else if(ext != VBComponentExtensions.FormBinaryExtension)
            {
                components.Import(filePath);
            }
        }
    }
}
