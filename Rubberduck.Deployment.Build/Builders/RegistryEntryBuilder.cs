using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Xml;
using Microsoft.Win32;
using Rubberduck.Deployment.Build.Structs;

namespace Rubberduck.Deployment.Build.Builders
{
    public class RegistryEntryBuilder
    {
        private FileMap fileMap;
        private TypeLibMap typeLibMap;
        private Dictionary<string, List<RegistryEntry>> interfaceMap;
        private List<ClassMap> classMapList;
        private Dictionary<string, List<RegistryEntry>> recordMap;

        private const string baseFolder = @"Software\Classes\";

        public IOrderedEnumerable<RegistryEntry> Parse(string tlbHeatOutputXmlPath, string dllHeatOutputXmlPath)
        {
            try
            {
                var tlbXml = new XmlDocument();
                tlbXml.Load(new XmlTextReader(tlbHeatOutputXmlPath) {Namespaces = false});

                var dllXml = new XmlDocument();
                dllXml.Load(new XmlTextReader(dllHeatOutputXmlPath) {Namespaces = false});

                fileMap = ExtractFilePath(dllXml.SelectSingleNode("//File"));
                typeLibMap = ExtractTypeLib(tlbXml.SelectSingleNode("//TypeLib"));
                interfaceMap = ExtractInterfaces(tlbXml.SelectNodes(@"//TypeLib/Interface"), typeLibMap);
                classMapList = ExtractClasses(dllXml.SelectSingleNode("//File"), typeLibMap);
                recordMap = ExtractRecords(dllXml.SelectNodes(@"//RegistryValue[starts-with(@Key, 'Record\')]"));

                var tmp = new List<RegistryEntry>();
                tmp.AddRange(typeLibMap.Entries);

                foreach (var map in interfaceMap)
                {
                    tmp.AddRange(map.Value);
                }

                foreach (var classMap in classMapList)
                {
                    tmp.AddRange(classMap.Entries);
                }

                foreach (var record in recordMap)
                {
                    tmp.AddRange(record.Value);
                }

                return tmp.OrderBy(t => t.Key);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        public FileMap ExtractFilePath(XmlNode node)
        {
            // <File Id="filEC54EFF475370C463E57905243620C96" KeyPath="yes" Source="SourceDir\Debug\Rubberduck.dll" />
            var fileNode = node.SelectSingleNode("//File");
            return new FileMap(fileNode.Attributes["Id"].Value,
                PlaceHolders.DllPath); 
                //fileNode.Attributes["Source"].Value);
        }

        public TypeLibMap ExtractTypeLib(XmlNode node)
        {
            /*
                <TypeLib Id="{E07C841C-14B4-4890-83E9-8C80B06DD59D}" Description="Rubberduck" HelpDirectory="dir39B22699688E51DCD8DCBB99A47E835B" Language="0" MajorVersion="2" MinorVersion="1">
                
                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}]

                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}\2.1]
                @="Rubberduck"

                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}\2.1\0]

                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}\2.1\0\win32]
                @="C:\\Github\\Rubberduck\\Rubberduck\\RetailCoder.VBE\\bin\\Debug\\Rubberduck.tlb"

                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}\2.1\FLAGS]
                @="0"

                [HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\{E07C841C-14B4-4890-83E9-8C80B06DD59D}\2.1\HELPDIR]
                @="C:\\Github\\Rubberduck\\Rubberduck\\RetailCoder.VBE\\bin\\Debug"
            */
            const string basePath = baseFolder + @"TypeLib\";
            var libGuid = node.Attributes["Id"].Value; //mandatory; it should throw if we don't have it.
            var libDesc = node.Attributes["Description"]?.Value ?? string.Empty;
            var language = node.Attributes["Language"]?.Value ?? "0";
            var major = node.Attributes["MajorVersion"]?.Value ?? "1";
            var minor = node.Attributes["MinorVersion"]?.Value ?? "0";
            var version = major + "." + minor;
            
            var flags = 0;
            {
                // WiX doesn't have a flags attribute but has individual attributes that maps to 
                // each LIBFLAG enumeration
                // https://msdn.microsoft.com/en-us/library/windows/desktop/ms221610(v=vs.85).aspx
                // https://msdn.microsoft.com/en-us/library/windows/desktop/ms221149(v=vs.85).aspx
                // http://wixtoolset.org/documentation/manual/v3/xsd/wix/typelib.html

                flags = flags | (node.Attributes["Control"]?.Value == "yes" ? (int) LIBFLAGS.LIBFLAG_FCONTROL : 0);
                flags = flags | (node.Attributes["HasDiskImage"]?.Value == "yes" ? (int) LIBFLAGS.LIBFLAG_FHASDISKIMAGE : 0);
                flags = flags | (node.Attributes["Hidden"]?.Value == "yes" ? (int) LIBFLAGS.LIBFLAG_FHIDDEN : 0);
                flags = flags | (node.Attributes["Restricted"]?.Value == "yes" ? (int) LIBFLAGS.LIBFLAG_FRESTRICTED : 0);
            }

            var entries = new List<RegistryEntry>
            {
                new RegistryEntry($@"{basePath}{libGuid}\", null, null, RegistryValueKind.None, Bitness.IsAgnostic, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}", null, libDesc, RegistryValueKind.String, Bitness.IsAgnostic, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}\{language}", null, null, RegistryValueKind.None, Bitness.IsAgnostic, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}\{language}\win32", null, PlaceHolders.TlbPath, RegistryValueKind.ExpandString, Bitness.Is32Bit, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}\{language}\win64", null, PlaceHolders.TlbPath, RegistryValueKind.ExpandString, Bitness.Is64Bit, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}\FLAGS", null, flags.ToString(), RegistryValueKind.String, Bitness.IsAgnostic, fileMap),
                new RegistryEntry($@"{basePath}{libGuid}\{version}\HELPDIR", null, PlaceHolders.InstallPath, RegistryValueKind.ExpandString, Bitness.IsAgnostic, fileMap)
            };

            return new TypeLibMap(libGuid, version, entries);
        }

        public Dictionary<string, List<RegistryEntry>> ExtractInterfaces(XmlNodeList interfaceList, TypeLibMap typeLibMap)
        {
            var dict = new Dictionary<string, List<RegistryEntry>>();
            foreach (XmlNode node in interfaceList)
            {
                /*
                    <Interface Id="{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}" Name="IParserState" ProxyStubClassId32="{00020424-0000-0000-C000-000000000046}" />                 

                    [HKEY_LOCAL_MACHINE\Software\Classes\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}]
                    @="IParserState"

                    [HKEY_LOCAL_MACHINE\Software\Classes\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}\ProxyStubClsid32]
                    @="{00020424-0000-0000-C000-000000000046}"

                    [HKEY_LOCAL_MACHINE\Software\Classes\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}\TypeLib]
                    @="{E07C841C-14B4-4890-83E9-8C80B06DD59D}"
                    "Version"="2.1"

                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}]
                    @="IParserState"

                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}\ProxyStubClsid32]
                    @="{00020424-0000-0000-C000-000000000046}"

                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\Interface\{02FA52F2-0D39-30DF-AB33-E8695C7E3A36}\TypeLib]
                    @="{E07C841C-14B4-4890-83E9-8C80B06DD59D}"
                    "Version"="2.1"
                */

                var guid = node.Attributes["Id"].Value;
                var name = node.Attributes["Name"].Value;
                var proxy = node.Attributes["ProxyStubClassId"]?.Value;
                var proxy32 = node.Attributes["ProxyStubClassId32"]?.Value;

                var entries = new List<RegistryEntry>();
                entries.Add(new RegistryEntry($@"{baseFolder}Interface\{guid}", null, name, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                if (proxy != null)
                {
                    entries.Add(new RegistryEntry($@"{baseFolder}Interface\{guid}\ProxyStubClsid", null, proxy,
                        RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                }
                if (proxy32 != null)
                {
                    entries.Add(new RegistryEntry($@"{baseFolder}Interface\{guid}\ProxyStubClsid32", null, proxy32,
                        RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                }
                entries.Add(new RegistryEntry($@"{baseFolder}Interface\{guid}\TypeLib", null, typeLibMap.Guid, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                entries.Add(new RegistryEntry($@"{baseFolder}Interface\{guid}\TypeLib", "Version", typeLibMap.Version, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));

                dict.Add(guid, entries);
            }

            return dict;
        }

        public Dictionary<string, List<RegistryEntry>> ExtractRecords(XmlNodeList recordList)
        {
            /*
                <RegistryValue Root="HKCR" Key="Record\{3E077C17-5678-3605-8449-FEABE42C9725}\2.1.6644.3188" Name="Class" Value="Rubberduck.API.VBA.DeclarationType" Type="string" Action="write" />
                <RegistryValue Root="HKCR" Key="Record\{3E077C17-5678-3605-8449-FEABE42C9725}\2.1.6644.3188" Name="Assembly" Value="Rubberduck, Version=2.1.6644.3188, Culture=neutral, PublicKeyToken=null" Type="string" Action="write" />
                <RegistryValue Root="HKCR" Key="Record\{3E077C17-5678-3605-8449-FEABE42C9725}\2.1.6644.3188" Name="RuntimeVersion" Value="v4.0.30319" Type="string" Action="write" />
                <RegistryValue Root="HKCR" Key="Record\{3E077C17-5678-3605-8449-FEABE42C9725}\2.1.6644.3188" Name="CodeBase" Value="file:///[#filEC54EFF475370C463E57905243620C96]" Type="string" Action="write" />
            */

            var dict = new Dictionary<string, List<RegistryEntry>>();
            foreach (XmlNode node in recordList)
            {
                var key = baseFolder + node.Attributes["Key"].Value;
                var name = node.Attributes["Name"].Value;
                var value = node.Attributes["Value"].Value;
                var type = ParseRegistryType(node.Attributes["Type"].Value);

                var guid = key.Split('\\')[1];
                if (!dict.ContainsKey(guid))
                {
                    dict.Add(guid, new List<RegistryEntry>());
                }

                dict[guid].Add(new RegistryEntry(key, name, value, type, Bitness.IsAgnostic, fileMap));
            }

            return dict;
        }

        public List<ClassMap> ExtractClasses(XmlNode componentXml, TypeLibMap typeLibMap)
        {
            var classMaps = new List<ClassMap>();

            foreach (XmlNode node in componentXml.SelectNodes("//Class"))
            {
                /*
                    <Class Id="{28754D11-10CC-45FD-9F6A-525A65412B7A}" Context="InprocServer32" Description="Rubberduck.API.VBA.ParserState" ThreadingModel="both" ForeignServer="mscoree.dll">
                        <ProgId Id="Rubberduck.ParserState" Description="Rubberduck.API.VBA.ParserState" />
                    </Class> 
                */
                var guid = node.Attributes["Id"].Value;
                var context = node.Attributes["Context"].Value;
                var foreignServer = node.Attributes["ForeignServer"].Value;
                var description = node.Attributes["Description"].Value;
                var threadingModel = node.Attributes["ThreadingModel"].Value;
                var progId = node.FirstChild.Attributes["Id"].Value;
                var progIdDescription = node.FirstChild.Attributes["Description"].Value;
                var entries = new List<RegistryEntry>();

                /*
                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}]
                    @="Rubberduck.API.VBA.ParserState"
                        
                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32]
                    @="mscoree.dll"
                    "ThreadingModel"="Both"
                        
                    [HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node\CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\ProgId]
                    @="Rubberduck.ParserState"                 
                */
                entries.Add(new RegistryEntry($@"{baseFolder}CLSID\{guid}", null, description, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                entries.Add(new RegistryEntry($@"{baseFolder}CLSID\{guid}\{context}", null, foreignServer, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                entries.Add(new RegistryEntry($@"{baseFolder}CLSID\{guid}\{context}", "ThreadingModel", threadingModel, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                entries.Add(new RegistryEntry($@"{baseFolder}CLSID\{guid}\ProgId", null, progId, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));

                foreach (XmlNode registryNode in componentXml.SelectNodes($@"//RegistryValue[starts-with(@Key, 'CLSID\{guid}')]"))
                {

                    /*
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" Value="" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32\2.1.6644.3188" Name="Class" Value="Rubberduck.API.VBA.ParserState" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32\2.1.6644.3188" Name="Assembly" Value="Rubberduck, Version=2.1.6644.3188, Culture=neutral, PublicKeyToken=null" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32\2.1.6644.3188" Name="RuntimeVersion" Value="v4.0.30319" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32\2.1.6644.3188" Name="CodeBase" Value="file:///[#filEC54EFF475370C463E57905243620C96]" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32" Name="Class" Value="Rubberduck.API.VBA.ParserState" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32" Name="Assembly" Value="Rubberduck, Version=2.1.6644.3188, Culture=neutral, PublicKeyToken=null" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32" Name="RuntimeVersion" Value="v4.0.30319" Type="string" Action="write" />
                        <RegistryValue Root="HKCR" Key="CLSID\{28754D11-10CC-45FD-9F6A-525A65412B7A}\InprocServer32" Name="CodeBase" Value="file:///[#filEC54EFF475370C463E57905243620C96]" Type="string" Action="write" />                
                    */

                    var key = baseFolder + registryNode.Attributes["Key"].Value;
                    var name = registryNode.Attributes["Name"]?.Value;
                    var value = registryNode.Attributes["Value"].Value;
                    var type = ParseRegistryType(registryNode.Attributes["Type"].Value);
                    entries.Add(new RegistryEntry(key, name, value, string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(value) ? RegistryValueKind.None : type, Bitness.IsPlatformDependent, fileMap));
                }

                {
                    /*
                        [HKEY_LOCAL_MACHINE\Software\Classes\Rubberduck.ParserState]
                        @="Rubberduck.API.VBA.ParserState"

                        [HKEY_LOCAL_MACHINE\Software\Classes\Rubberduck.ParserState\CLSID]
                        @="{28754D11-10CC-45FD-9F6A-525A65412B7A}" 
                    */

                    entries.Add(new RegistryEntry(baseFolder + progId, null, progIdDescription, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                    entries.Add(new RegistryEntry(baseFolder + progId + @"\CLSID", null, guid, RegistryValueKind.String, Bitness.IsPlatformDependent, fileMap));
                }

                classMaps.Add(new ClassMap(guid, context, description, threadingModel, progId, progIdDescription, entries));
            }

            return classMaps;
        }
        
        private RegistryValueKind ParseRegistryType(string type)
        {
            // http://wixtoolset.org/documentation/manual/v3/xsd/wix/registryvalue.html
            switch (type)
            {
                case "string":
                    return RegistryValueKind.String;
                case "integer":
                    return RegistryValueKind.DWord;
                case "binary":
                    return RegistryValueKind.Binary;
                case "expandable":
                    return RegistryValueKind.ExpandString;
                case "multiString":
                    return RegistryValueKind.MultiString;
                default:
                    return RegistryValueKind.None;
            }
        }
    }
}

