<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:wix="http://schemas.microsoft.com/wix/2006/wi" >
  <!-- The name of the project reference -->
  <xsl:param name="projectName" select="''"></xsl:param>
  <!-- The path of the project reference -->
  <xsl:param name="projectFilePath" select="''"></xsl:param>
  <!-- The intermediate directory of the setup project -->
  <xsl:param name="intermediateDir" select="''"></xsl:param>

  <xsl:output method='xml' indent='yes' cdata-section-elements='wix:Condition'/>

  <!-- The content of the XML file for the project output -->
  <xsl:variable name="outputSettings" select="document(concat('_', $projectName, '.xml'))"/>
  <!-- Select all 'COMEntry' elements in the 'RegisterForCOM.xml' file -->
  <xsl:variable name="comEntries" select="document('RegisterForCOM.xml')"/>

  <xsl:variable name='newline'>
    <xsl:text>&#13;&#10;</xsl:text>
  </xsl:variable>

  <xsl:variable name='urlspace'>
    <xsl:text>%20</xsl:text>
  </xsl:variable>

  <xsl:variable name='justspace'>
    <xsl:text><![CDATA[ ]]></xsl:text>
  </xsl:variable>

  <xsl:variable name='file-extension-delimiter'>
    <xsl:text><![CDATA[.]]></xsl:text>
  </xsl:variable>

  <!-- The list of attributes supported in the 'File' element of WiX -->
  <xsl:variable name="file-attribute-list">
    <Item>Assembly</Item>
    <Item>AssemblyApplication</Item>
    <Item>AssemblyManifest</Item>
    <Item>BindPath</Item>
    <Item>Checksum</Item>
    <Item>CompanionFile</Item>
    <Item>Compressed</Item>
    <Item>DefaultLanguage</Item>
    <Item>DefaultSize</Item>
    <Item>DefaultVersion</Item>
    <Item>DiskId</Item>
    <Item>FontTitle</Item>
    <Item>Hidden</Item>
    <Item>KeyPath</Item>
    <Item>LongName</Item>
    <Item>Name</Item>
    <Item>PatchAllowIgnoreOnError</Item>
    <Item>PatchGroup</Item>
    <Item>PatchIgnore</Item>
    <Item>PatchWholeFile</Item>
    <Item>ProcessorArchitecture</Item>
    <Item>ReadOnly</Item>
    <Item>SelfRegCost</Item>
    <Item>ShortName</Item>
    <Item>System</Item>
    <Item>TrueType</Item>
    <Item>Vital</Item>
  </xsl:variable>

  <!-- The list of attributes supported in the 'Component' element of WiX -->
  <xsl:variable name="component-attribute-list">
    <Item>Permanent</Item>
    <Item>SharedDllRefCount</Item>
    <Item>Transitive</Item>
  </xsl:variable>

  <!-- Copy nodes and attributes -->
  <xsl:template match="@*|node()">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()" />
    </xsl:copy>
  </xsl:template>

  <xsl:template name="substring-before-last">
    <xsl:param name="input" />
    <xsl:param name="substr" />
    <xsl:if test="$substr and contains($input, $substr)">
      <xsl:variable name="temp" select="substring-after($input, $substr)" />
      <xsl:value-of select="substring-before($input, $substr)" />
      <xsl:if test="contains($temp, $substr)">
        <xsl:value-of select="$substr" />
        <xsl:call-template name="substring-before-last">
          <xsl:with-param name="input" select="$temp" />
          <xsl:with-param name="substr" select="$substr" />
        </xsl:call-template>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <xsl:template name="substring-after-last">
    <xsl:param name="input"/>
    <xsl:param name="substr"/>
    <!-- Extract the string which comes after the first occurrence -->
    <xsl:variable name="temp" select="substring-after($input,$substr)"/>
    <xsl:choose>
      <!-- If it still contains the search string the recursively process -->
      <xsl:when test="$substr and contains($temp,$substr)">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="input" select="$temp"/>
          <xsl:with-param name="substr" select="$substr"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$temp"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Extract the file name without extension from the full path -->
  <xsl:template name="fileName-without-extension">
    <xsl:param name="path" />
    <xsl:choose>
      <xsl:when test="contains($path,'\')">
        <xsl:call-template name="fileName-without-extension">
          <xsl:with-param name="path" select="substring-after($path,'\')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($path,'/')">
        <xsl:call-template name="fileName-without-extension">
          <xsl:with-param name="path" select="substring-after($path,'/')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($path, $file-extension-delimiter)">
        <xsl:call-template name="substring-before-last">
          <xsl:with-param name="input" select="$path" />
          <xsl:with-param name="substr" select="$file-extension-delimiter" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$path" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Extract the file extension -->
  <xsl:template name="file-extension">
    <xsl:param name="path" />
    <xsl:variable name="fileName">
      <xsl:call-template name="fileName-with-extension">
        <xsl:with-param name="path" select="$path" />
      </xsl:call-template>
    </xsl:variable>
    <xsl:call-template name="substring-after-last">
      <xsl:with-param name="input" select="$fileName" />
      <xsl:with-param name="substr" select="$file-extension-delimiter" />
    </xsl:call-template>
  </xsl:template>

  <!-- Extract the file name from the full path -->
  <xsl:template name="fileName-with-extension">
    <xsl:param name="path" />
    <xsl:choose>
      <xsl:when test="contains($path,'\')">
        <xsl:call-template name="fileName-with-extension">
          <xsl:with-param name="path" select="substring-after($path,'\')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($path,'/')">
        <xsl:call-template name="fileName-with-extension">
          <xsl:with-param name="path" select="substring-after($path,'/')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$path" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Remove double slashes from the path and replace '/' with '\' -->
  <xsl:template name="normalize-dirname">
    <xsl:param name="path"/>
    <xsl:variable name="path-" select="translate($path, '/', '\')"/>
    <xsl:choose>
      <xsl:when test="contains($path-, '\\')">
        <xsl:variable name="pa" select="substring-before($path-, '\\')"/>
        <xsl:variable name="th" select="substring-after($path-, '\\')"/>
        <xsl:variable name="pa-th" select="concat($pa, '\', $th)"/>
        <xsl:call-template name="normalize-dirname">
          <xsl:with-param name="path" select="$pa-th"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$path-"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Create an array from the 'file-attribute-list' variable -->
  <xsl:param name="file-attributes" select="document('')/*/xsl:variable[@name='file-attribute-list']/*"/>
  <!-- Create an array from the 'component-attribute-list' variable -->
  <xsl:param name="component-attributes" select="document('')/*/xsl:variable[@name='component-attribute-list']/*"/>

  <!-- Save the name of the project from the full path to the project reference -->
  <xsl:variable name="projectFileNameWithoutExtension">
    <xsl:call-template name="fileName-without-extension">
      <xsl:with-param name="path" select="$projectFilePath"/>
    </xsl:call-template>
  </xsl:variable>

  <!-- Return a new string in which all occurrences of a specified text are replaced with another specified text -->
  <xsl:template name="replace-string">
    <!-- The original text -->
    <xsl:param name="text"/>
    <!-- The text to be replaced -->
    <xsl:param name="replace"/>
    <!-- The text to replace all occurrences of the 'replace' parameter -->
    <xsl:param name="with"/>
    <xsl:choose>
      <xsl:when test="contains($text,$replace)">
        <xsl:value-of select="substring-before($text,$replace)"/>
        <xsl:value-of select="$with"/>
        <xsl:call-template name="replace-string">
          <xsl:with-param name="text" select="substring-after($text,$replace)"/>
          <xsl:with-param name="replace" select="$replace"/>
          <xsl:with-param name="with" select="$with"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$text"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="replace-project-variables">
    <!-- The source path -->
    <xsl:param name="path"/>
    <!-- The project directory -->
    <xsl:param name="projectDirVar"/>
    <!-- The project target directory -->
    <xsl:param name="targetDirVar"/>
    <!-- The original project name -->
    <xsl:param name="projectReferenceName"/>
    <xsl:variable name="normalizedPath">
      <xsl:call-template name="replace-string">
        <xsl:with-param name="text" select="$path"/>
        <xsl:with-param name="replace" select="$urlspace"/>
        <xsl:with-param name="with" select="$justspace" />
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="var_ProjectDir" select="concat('$(var.', $projectReferenceName, '.ProjectDir)')"/>
    <xsl:variable name="var_TargetDir" select="concat('$(var.', $projectReferenceName, '.TargetDir)')"/>
    <xsl:choose>
      <xsl:when test="starts-with($normalizedPath, $var_ProjectDir)">
        <xsl:call-template name="replace-string">
          <xsl:with-param name="text" select="$normalizedPath"/>
          <xsl:with-param name="replace" select="$var_ProjectDir"/>
          <xsl:with-param name="with" select="$projectDirVar"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="starts-with($normalizedPath, $var_TargetDir)">
        <xsl:call-template name="replace-string">
          <xsl:with-param name="text" select="$normalizedPath"/>
          <xsl:with-param name="replace" select="$var_TargetDir"/>
          <xsl:with-param name="with" select="$targetDirVar"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$normalizedPath"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Copy an element recursively. -->
  <xsl:template name="copy-all-nodes">
    <!-- The element to copy -->
    <xsl:param name="sourceElem"/>
    <!-- The share file id -->
    <xsl:param name="fileId"/>
    <xsl:element name="{name($sourceElem)}" namespace="http://schemas.microsoft.com/wix/2006/wi">
      <xsl:for-each select="@*">
        <xsl:variable name="attrValue" select="."/>
        <xsl:choose>
          <xsl:when test="contains($attrValue, '[#')">
            <xsl:attribute name="{name(.)}">
              <xsl:value-of select="substring-before($attrValue, '[#')"/>
              <xsl:text>[#</xsl:text>
              <xsl:value-of select="$fileId"/>
              <xsl:text>]</xsl:text>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      <xsl:for-each select="$sourceElem//wix:*[parent::* = $sourceElem]">
        <xsl:call-template name="copy-all-nodes">
          <xsl:with-param name="sourceElem" select="."/>
          <xsl:with-param name="fileId" select="$fileId"/>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:element>
  </xsl:template>

  <!-- Apply template for 'File' elements -->
  <xsl:template match="wix:File">
    <xsl:copy>
      <!-- Get the share file id from the XML file of the project reference -->
      <xsl:variable name="shareFileId">
        <xsl:value-of select="$outputSettings//ProjectOutput/@SharedFileId"/>
      </xsl:variable>
      <xsl:variable name="tag">
        <xsl:value-of select="$outputSettings//ProjectOutput/@Tag"/>
      </xsl:variable>
      <xsl:variable name="sourceExtension">
        <xsl:call-template name="file-extension">
          <xsl:with-param name="path" select="@Source"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="groupType">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="input" select="ancestor::wix:DirectoryRef/@Id" />
          <xsl:with-param name="substr" select="$file-extension-delimiter" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <!-- Change the id for the .exe and .dll file -->
        <xsl:when test="$shareFileId != '' and $groupType = 'Binaries' and (translate($sourceExtension,'EXE','exe') = 'exe' or translate($sourceExtension,'DLL','dll') = 'dll')">
          <xsl:attribute name="Id">
            <xsl:value-of select="$shareFileId"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$tag != ''">
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
            <xsl:value-of select="$tag"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <!-- Create attributes from the 'file-attributes' array -->
      <!-- All values are loaded from the XML file of the project reference -->
      <xsl:for-each select="$file-attributes">
        <xsl:variable name="attrName" select="."/>
        <xsl:for-each select="$outputSettings//*[parent::*[name()=$groupType]]">
          <xsl:if test="(normalize-space(.) != '') and ($attrName = name(.))">
            <xsl:attribute name="{$attrName}">
              <xsl:value-of select="."/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
      <!-- Create the 'Source' attribute and replace the name of the project reference with the actual project name without the '_' suffix) -->
      <xsl:attribute name="Source">
        <xsl:call-template name="replace-project-variables">
          <xsl:with-param name="path" select="@Source"/>
          <xsl:with-param name="projectDirVar" select="$outputSettings//ProjectOutput/@ProjectDirVar" />
          <xsl:with-param name="targetDirVar" select="$outputSettings//ProjectOutput/@TargetDirVar" />
          <xsl:with-param name="projectReferenceName" select="$outputSettings//ProjectOutput/@Name" />
        </xsl:call-template>
      </xsl:attribute>
      <xsl:if test="$groupType = 'Binaries'">
        <!-- Search for the 'COMEntry' node in the 'RegisterForCOM.xml' file -->
        <xsl:variable name="comEntry" select="$comEntries//COMEntry[@Name=$projectName]"/>
        <!-- Copy all registry entries if they exist in the '<project reference>.com.wsx' file -->
        <xsl:if test="$comEntry">
          <xsl:variable name="comFileName">
            <xsl:text>_</xsl:text>
            <xsl:call-template name="fileName-without-extension">
              <xsl:with-param name="path" select="$comEntry/@Source"/>
            </xsl:call-template>
            <xsl:text>.com.wsx</xsl:text>
          </xsl:variable>
          <xsl:variable name="comFilePath">
            <xsl:call-template name="normalize-dirname">
              <xsl:with-param name="path" select="concat('..\', $intermediateDir, $comFileName)"/>
            </xsl:call-template>
          </xsl:variable>
          <!-- We copy all child nodes of the 'File' element only -->
          <xsl:for-each select="document($comFilePath)//wix:File//wix:*[parent::*[name()='File']]">
            <xsl:value-of select="$newline"/>
            <xsl:copy-of select="."/>
            <xsl:value-of select="$newline"/>
          </xsl:for-each>
        </xsl:if>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <!-- Apply template for 'Component' elements -->
  <xsl:template match="wix:Component">
    <xsl:copy>
      <!-- Get the share file id from the XML file of the project reference -->
      <xsl:variable name="shareFileId">
        <xsl:value-of select="$outputSettings//ProjectOutput/@SharedFileId"/>
      </xsl:variable>
      <xsl:variable name="tag">
        <xsl:value-of select="$outputSettings//ProjectOutput/@Tag"/>
      </xsl:variable>
      <xsl:variable name="sourceExtension">
        <xsl:call-template name="file-extension">
          <xsl:with-param name="path" select=".//wix:File/@Source"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="groupType">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="input" select="ancestor::wix:DirectoryRef/@Id" />
          <xsl:with-param name="substr" select="$file-extension-delimiter" />
        </xsl:call-template>
      </xsl:variable>
      <!-- The 'Condition' element from the project output XML -->
      <xsl:variable name="condition" select="normalize-space($outputSettings//Condition[parent::*[name()=$groupType]])"/>
      <!-- The HKCU key from the project output XML -->
      <xsl:variable name="hkcuKey" select="normalize-space($outputSettings//HKCUKey[parent::*[name()=$groupType]])"/>
      <!-- The 'FileTypes' element from the project output XML -->
      <xsl:variable name="fileTypes" select="$outputSettings//FileTypes[parent::*[name()=$groupType]]"/>
      <xsl:choose>
        <!-- Change the id for the .exe and .dll file -->
        <xsl:when test="$shareFileId != '' and $groupType = 'Binaries' and (translate($sourceExtension,'EXE','exe') = 'exe' or translate($sourceExtension,'DLL','dll') = 'dll')">
          <xsl:attribute name="Id">
            <xsl:text>com</xsl:text>
            <xsl:value-of select="$shareFileId"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$tag != ''">
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
            <xsl:value-of select="$tag"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name="Guid">
        <xsl:value-of select="@Guid"/>
      </xsl:attribute>
      <xsl:variable name="defaultComponentId" select="@Id" />
      <!-- Create attributes from the 'component-attributes' array -->
      <!-- All values are loaded from the XML file of the project reference -->
      <xsl:for-each select="$component-attributes">
        <xsl:variable name="attrName" select="."/>
        <xsl:for-each select="$outputSettings//*[parent::*[name()=$groupType]]">
          <xsl:if test="(normalize-space(.) != '') and ($attrName = name(.))">
            <xsl:attribute name="{$attrName}">
              <xsl:value-of select="."/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:if test="$groupType = 'Binaries'">
        <!-- Search for the 'COMEntry' element for the project output in the 'RegisterForCOM.xml' file -->
        <xsl:variable name="comEntry" select="$comEntries//COMEntry[@Name=$projectName]"/>
        <!-- If the project output is registered for COM, we load the registry information from the '<project name>.com.wsx' file -->
        <xsl:if test="$comEntry">
          <xsl:variable name="comFileName">
            <xsl:text>_</xsl:text>
            <xsl:call-template name="fileName-without-extension">
              <xsl:with-param name="path" select="$comEntry/@Source"/>
            </xsl:call-template>
            <xsl:text>.com.wsx</xsl:text>
          </xsl:variable>
          <xsl:variable name="comFilePath">
            <xsl:call-template name="normalize-dirname">
              <xsl:with-param name="path" select="concat('..\', $intermediateDir, $comFileName)"/>
            </xsl:call-template>
          </xsl:variable>
          <!-- We copy all nodes except for 'File' and 'TypeLib' ('File' and 'TypeLib' are copied by the code in the File template)-->
          <xsl:for-each select="document($comFilePath)//wix:Component//wix:*[parent::*[name()='Component']]">
            <xsl:if test="(name(.)!='File') and (name(.)!='TypeLib')">
              <xsl:value-of select="$newline"/>
              <!-- Create a new element and copy all child nodes -->
              <xsl:call-template name="copy-all-nodes">
                <xsl:with-param name="sourceElem" select="."/>
                <xsl:with-param name="fileId" select="$shareFileId"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:if>
      <xsl:apply-templates/>
      <xsl:if test="$condition != ''">
        <!-- Create the 'Condition' element -->
        <xsl:element name="Condition" namespace="http://schemas.microsoft.com/wix/2006/wi">
          <xsl:value-of select="$condition"/>
        </xsl:element>
        <xsl:value-of select="$newline"/>
      </xsl:if>
      <xsl:if test="$hkcuKey != ''">
        <!-- Create the 'RegistryValue' element -->
        <xsl:element name="RegistryValue" namespace="http://schemas.microsoft.com/wix/2006/wi">
          <xsl:attribute name="Root">
            <xsl:text>HKCU</xsl:text>
          </xsl:attribute>
          <xsl:attribute name="Key">
            <xsl:value-of select="$hkcuKey"/>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="$shareFileId != '' and $groupType = 'Binaries' and (translate($sourceExtension,'EXE','exe') = 'exe' or translate($sourceExtension,'DLL','dll') = 'dll')">
              <xsl:attribute name="Name">
                <xsl:text>com</xsl:text>
                <xsl:value-of select="$shareFileId"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="Name">
                <xsl:value-of select="$defaultComponentId"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:attribute name="Type">
            <xsl:text>string</xsl:text>
          </xsl:attribute>
          <xsl:attribute name="Value">
            <xsl:call-template name="fileName-with-extension">
              <xsl:with-param name="path" select=".//wix:File/@Source"/>
            </xsl:call-template>
            <xsl:text> project output file</xsl:text>
          </xsl:attribute>
          <xsl:attribute name="KeyPath">
            <xsl:text>yes</xsl:text>
          </xsl:attribute>
        </xsl:element>
        <xsl:value-of select="$newline"/>
      </xsl:if>
      <xsl:if test="count($fileTypes//wix:*) > 0 and $shareFileId != '' and $groupType = 'Binaries' and (translate($sourceExtension,'EXE','exe') = 'exe' or translate($sourceExtension,'DLL','dll') = 'dll')">
        <!-- Create the 'ProgId' elements -->
        <xsl:for-each select="$fileTypes//wix:*[parent::*[name()='FileTypes']]" >
          <xsl:copy-of select="."/>
        </xsl:for-each>
        <xsl:value-of select="$newline"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="wix:Directory">
    <xsl:copy>
      <xsl:variable name="tag">
        <xsl:value-of select="$outputSettings//ProjectOutput/@Tag"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$tag != ''">
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
            <xsl:value-of select="$tag"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="Id">
            <xsl:value-of select="@Id"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name="Name">
        <xsl:value-of select="@Name"/>
      </xsl:attribute>
      <xsl:variable name="dirId" >
        <xsl:value-of select="@Id"/>
      </xsl:variable>
      <xsl:variable name="componentId" >
        <xsl:text>com_</xsl:text>
        <xsl:value-of select="@Id"/>
      </xsl:variable>
      <xsl:variable name="groupType">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="input" select="ancestor::wix:DirectoryRef/@Id" />
          <xsl:with-param name="substr" select="$file-extension-delimiter" />
        </xsl:call-template>
      </xsl:variable>
      <!-- The HKCU key from the project output XML -->
      <xsl:variable name="hkcuKey" select="normalize-space($outputSettings//HKCUKey[parent::*[name()=$groupType]])"/>
      <xsl:if test="$hkcuKey != ''">
        <xsl:value-of select="$newline"/>
        <!-- Create the 'Component' element -->
        <xsl:element name="Component" namespace="http://schemas.microsoft.com/wix/2006/wi">
          <xsl:attribute name="Id">
            <xsl:value-of select="$componentId"/>
          </xsl:attribute>
          <xsl:value-of select="$newline"/>
          <!-- Create the 'RemoveFolder' element -->
          <xsl:element name="RemoveFolder" namespace="http://schemas.microsoft.com/wix/2006/wi">
            <xsl:attribute name="Id">
              <xsl:value-of select="$componentId"/>
            </xsl:attribute>
            <xsl:attribute name="On">
              <xsl:text>uninstall</xsl:text>
            </xsl:attribute>
          </xsl:element>
          <xsl:value-of select="$newline"/>
          <!-- Create the 'RegistryValue' element -->
          <xsl:element name="RegistryValue" namespace="http://schemas.microsoft.com/wix/2006/wi">
            <xsl:attribute name="Root">
              <xsl:text>HKCU</xsl:text>
            </xsl:attribute>
            <xsl:attribute name="Key">
              <xsl:value-of select="$hkcuKey"/>
            </xsl:attribute>
            <xsl:attribute name="Name">
              <xsl:value-of select="$componentId"/>
            </xsl:attribute>
            <xsl:attribute name="Type">
              <xsl:text>string</xsl:text>
            </xsl:attribute>
            <xsl:attribute name="Value">
              <xsl:value-of select="$dirId"/>
              <xsl:text> project output directory</xsl:text>
            </xsl:attribute>
            <xsl:attribute name="KeyPath">
              <xsl:text>yes</xsl:text>
            </xsl:attribute>
          </xsl:element>
          <xsl:value-of select="$newline"/>
        </xsl:element>
      </xsl:if>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="wix:ComponentGroup">
    <xsl:copy>
      <xsl:attribute name="Id">
        <xsl:value-of select="@Id"/>
      </xsl:attribute>
      <xsl:variable name="groupId">
        <xsl:value-of select="@Id"/>
      </xsl:variable>
      <xsl:variable name="groupType">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="input" select="$groupId" />
          <xsl:with-param name="substr" select="$file-extension-delimiter" />
        </xsl:call-template>
      </xsl:variable>
      <!-- The HKCU key from the project output XML -->
      <xsl:variable name="hkcuKey" select="normalize-space($outputSettings//HKCUKey[parent::*[name()=$groupType]])"/>
      <xsl:if test="$hkcuKey != ''">
        <!-- Add references for dynamically added directory components -->
        <xsl:for-each select="//wix:Directory[ancestor::wix:DirectoryRef[@Id=$groupId]]" >
          <xsl:value-of select="$newline"/>
          <xsl:element name="ComponentRef" namespace="http://schemas.microsoft.com/wix/2006/wi">
            <xsl:attribute name="Id">
              <xsl:text>com_</xsl:text>
              <xsl:value-of select="./@Id"/>
            </xsl:attribute>
          </xsl:element>
        </xsl:for-each>
      </xsl:if>
      <!-- Get the share file id from the XML file of the project reference -->
      <xsl:variable name="shareFileId">
        <xsl:value-of select="$outputSettings//ProjectOutput/@SharedFileId"/>
      </xsl:variable>
      <xsl:variable name="tag">
        <xsl:value-of select="$outputSettings//ProjectOutput/@Tag"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$shareFileId != '' and $groupType = 'Binaries'">
          <xsl:for-each select="//wix:ComponentRef[parent::wix:ComponentGroup[@Id=$groupId]]" >
            <xsl:variable name="refId" >
              <xsl:value-of select="@Id"/>
            </xsl:variable>
            <xsl:variable name="sourceExtension">
              <xsl:call-template name="file-extension">
                <xsl:with-param name="path" select="//wix:Component[@Id = $refId]//wix:File/@Source"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:choose>
              <!-- Change the id for the .exe and .dll file -->
              <xsl:when test="translate($sourceExtension,'EXE','exe') = 'exe' or translate($sourceExtension,'DLL','dll') = 'dll'" >
                <xsl:value-of select="$newline"/>
                <xsl:element name="ComponentRef" namespace="http://schemas.microsoft.com/wix/2006/wi">
                  <xsl:attribute name="Id">
                    <xsl:text>com</xsl:text>
                    <xsl:value-of select="$shareFileId"/>
                  </xsl:attribute>
                </xsl:element>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$newline"/>
                <xsl:choose>
                  <xsl:when test="$tag != ''">
                    <xsl:element name="ComponentRef" namespace="http://schemas.microsoft.com/wix/2006/wi">
                      <xsl:attribute name="Id">
                        <xsl:value-of select="./@Id"/>
                        <xsl:value-of select="$tag"/>
                      </xsl:attribute>
                    </xsl:element>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:element name="ComponentRef" namespace="http://schemas.microsoft.com/wix/2006/wi">
                      <xsl:attribute name="Id">
                        <xsl:value-of select="./@Id"/>
                      </xsl:attribute>
                    </xsl:element>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
          <xsl:value-of select="$newline"/>
        </xsl:when>
        <xsl:when test="$tag != ''">
          <xsl:for-each select="//wix:ComponentRef[parent::wix:ComponentGroup[@Id=$groupId]]" >
            <xsl:value-of select="$newline"/>
            <xsl:element name="ComponentRef" namespace="http://schemas.microsoft.com/wix/2006/wi">
              <xsl:attribute name="Id">
                <xsl:value-of select="./@Id"/>
                <xsl:value-of select="$tag"/>
              </xsl:attribute>
            </xsl:element>
          </xsl:for-each>
          <xsl:value-of select="$newline"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>