<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:wix="http://schemas.microsoft.com/wix/2006/wi" >
  <!-- The path to the COM file -->
  <xsl:param name="sourceFilePath" select="''"></xsl:param>

  <xsl:output method='xml' indent='yes' cdata-section-elements='wix:Condition'/>

  <!-- Select all 'COMEntry' elements in the 'RegisterForCOM.xml' file -->
  <xsl:variable name="comEntries" select="document('RegisterForCOM.xml')//COMLibraries"/>

  <!-- Copy nodes and attributes -->
  <xsl:template match="@*|node()">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()" />
    </xsl:copy>
  </xsl:template>

  <!-- Insert a new line -->
  <xsl:variable name='newline'>
    <xsl:text>&#13;&#10;</xsl:text>
  </xsl:variable>

  <xsl:variable name='file-extension-delimiter'>
    <xsl:text><![CDATA[.]]></xsl:text>
  </xsl:variable>

  <xsl:variable name='comfile-extension'>
    <xsl:text><![CDATA[.com]]></xsl:text>
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
    <Item>Condition</Item>
    <Item>HKCUKey</Item>
    <Item>FileTypes</Item>
  </xsl:variable>

  <!-- Create an array from the 'file-attribute-list' variable -->
  <xsl:param name="file-attributes" select="document('')/*/xsl:variable[@name='file-attribute-list']/*"/>
  <!-- Create an array from the 'component-attribute-list' variable -->
  <xsl:param name="component-attributes" select="document('')/*/xsl:variable[@name='component-attribute-list']/*"/>

  <!-- Extract file name with extension from the specified path -->
  <xsl:template name="filename-with-extension">
    <!-- The path to a file -->
    <xsl:param name="path"/>
    <xsl:variable name="path-" select="translate($path, '/', '\')"/>
    <xsl:choose>
      <xsl:when test="contains($path-, '\')">
        <xsl:call-template name="filename-with-extension">
          <xsl:with-param name="path" select="substring-after($path-, '\')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$path-"/>
      </xsl:otherwise>
    </xsl:choose>
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
  <xsl:template name="filename-without-extension">
    <xsl:param name="path" />
    <xsl:choose>
      <xsl:when test="contains($path,'\')">
        <xsl:call-template name="filename-without-extension">
          <xsl:with-param name="path" select="substring-after($path,'\')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($path,'/')">
        <xsl:call-template name="filename-without-extension">
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

  <!-- Extract the file name from the first position in the specified list (';' is separator) -->
  <xsl:template name="get-first-file-ref">
    <!-- The file Id from the 'RegisterForCOM.xml' file -->
    <xsl:param name="refIdList" />
    <!-- The file Id from the original 'Component' element -->
    <xsl:param name="defaultRefId" />
    <xsl:choose>
      <xsl:when test="$refIdList!='' and contains($refIdList, ';')">
        <xsl:value-of select="substring-before($refIdList, ';')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$defaultRefId"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

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

  <!-- Copy an element recursively and replace the file references. -->
  <xsl:template name="copy-and-replace">
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
        <xsl:call-template name="copy-and-replace">
          <xsl:with-param name="sourceElem" select="."/>
          <xsl:with-param name="fileId" select="$fileId"/>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:element>
  </xsl:template>

  <xsl:template name="copy-all-attributes">
    <xsl:param name="selection"/>
    <xsl:for-each select="$selection">
      <xsl:variable name="attrValue" select="."/>
      <xsl:attribute name="{name(.)}">
        <xsl:value-of select="$attrValue"/>
      </xsl:attribute>
    </xsl:for-each>
  </xsl:template>

  <!-- Create the 'TypeLib' element. -->
  <xsl:template name="create-typelib">
    <!-- The original 'TypeLib' node -->
    <xsl:param name="typeLibOriginal"/>
    <!-- The value of the 'DirectoryId' attribute of the 'COMEntry' element in the 'RegisterForCOM.xml' file -->
    <xsl:param name="directoryId"/>
    <!-- Create the 'TypeLib' element -->
    <!-- The share file id -->
    <xsl:param name="fileId"/>
    <xsl:element name="TypeLib" namespace="http://schemas.microsoft.com/wix/2006/wi">
      <!-- Copy all attributes from the original 'TypeLib' and change the value of the 'HelpDirectory' attribute if it exists -->
      <!-- We can copy all attributes at once: <xsl:copy-of select="..//wix:TypeLib/@*"/> -->
      <xsl:for-each select="$typeLibOriginal/@*">
        <xsl:variable name="attrValue" select="."/>
        <xsl:choose>
          <xsl:when test="(name(.) = 'HelpDirectory')">
            <xsl:attribute name="{name(.)}">
              <xsl:value-of select="$directoryId"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="contains($attrValue, '[#') and ($fileId != '')">
            <xsl:attribute name="{name(.)}">
              <xsl:value-of select="substring-before($attrValue, '[#')"/>
              <xsl:text>[#</xsl:text>
              <xsl:value-of select="$fileId"/>
              <xsl:text>]</xsl:text>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="$attrValue != ''">
              <xsl:copy-of select="."/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      <!-- Add the 'Language' attribute if it doesn't exist -->
      <xsl:if test="not (./@Language)">
        <xsl:attribute name="Language">
          <xsl:value-of select="0"/>
        </xsl:attribute>
      </xsl:if>
      <!-- Copy all child nodes of the original 'TypeLib' element -->
      <!-- We can copy all elements at once: <xsl:copy-of select="..//wix:TypeLib//*"/> -->
      <xsl:for-each select="$typeLibOriginal//*[parent::*[name()='TypeLib']]">
        <xsl:value-of select="$newline"/>
        <xsl:choose>
          <xsl:when test="$fileId != ''">
            <xsl:call-template name="copy-and-replace">
              <xsl:with-param name="sourceElem" select="."/>
              <xsl:with-param name="fileId" select="$fileId"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      <xsl:value-of select="$newline"/>
    </xsl:element>
  </xsl:template>

  <!-- Copy the DirectoryRef element-->
  <xsl:template match="wix:DirectoryRef">
    <xsl:copy>
      <xsl:variable name="fileName">
        <xsl:call-template name="filename-without-extension">
          <xsl:with-param name="path" select="$sourceFilePath"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- Save the entry id -->
      <xsl:variable name="entryId">
        <xsl:call-template name="substring-before-last">
          <xsl:with-param name="input" select="$fileName"/>
          <xsl:with-param name="substr" select="$comfile-extension"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- Search for the 'COMEntry' node in the 'RegisterForCOM.xml' file -->
      <xsl:variable name="comFileNode" select="$comEntries//COMEntry[@Id=$entryId]"/>
      <!-- Determine the directory id -->
      <xsl:variable name="directoryId">
        <xsl:value-of select="$comFileNode/@DirectoryId"/>
      </xsl:variable>
      <!-- Create the 'Id' attribute -->
      <xsl:attribute name="Id">
        <xsl:value-of select="$directoryId"/>
      </xsl:attribute>
      <xsl:value-of select="$newline"/>
      <!-- Go to the 'Component' template -->
      <xsl:apply-templates select="wix:Component"/>
    </xsl:copy>
  </xsl:template>

  <!-- Copy the 'Component' element-->
  <xsl:template match="wix:Component">
    <xsl:copy>
      <xsl:variable name="fileName">
        <xsl:call-template name="filename-without-extension">
          <xsl:with-param name="path" select="$sourceFilePath"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- Save the entry id -->
      <xsl:variable name="entryId">
        <xsl:call-template name="substring-before-last">
          <xsl:with-param name="input" select="$fileName"/>
          <xsl:with-param name="substr" select="$comfile-extension"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- Search for the 'COMEntry' node in the 'RegisterForCOM.xml' file -->
      <xsl:variable name="comFileNode" select="$comEntries//COMEntry[@Id=$entryId]"/>
      <!-- Save the directory id -->
      <xsl:variable name="directoryId">
        <xsl:value-of select="$comFileNode/@DirectoryId"/>
      </xsl:variable>
      <!-- Save the COM file location -->
      <xsl:variable name="fileSource">
        <xsl:value-of select="$comFileNode/@Source"/>
      </xsl:variable>
      <!-- Save the share id of the file if it exists -->
      <xsl:variable name="shareFileId">
        <xsl:value-of select="$comFileNode/@SharedFileId"/>
      </xsl:variable>
      <xsl:if test="$comFileNode">
        <!-- Save the component id -->
        <xsl:variable name="componentId">
          <xsl:text>com</xsl:text>
          <xsl:value-of select="$shareFileId"/>
        </xsl:variable>
        <!-- Copy all attributes -->
        <xsl:for-each select="./@*">
          <xsl:choose>
            <xsl:when test="name(.) = 'Id'">
              <xsl:attribute name="Id">
                <xsl:value-of select="$componentId"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="{name(.)}">
                <xsl:value-of select="."/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
        <!-- Create attributes from the 'component-attributes' array -->
        <!-- All values are loaded from the 'COMEntry' node located in the 'RegisterForCOM.xml' file -->
        <xsl:for-each select="$component-attributes">
          <xsl:variable name="propertyName" select="."/>
          <xsl:variable name="propertyValue" select="normalize-space($comFileNode//*[name()=$propertyName])"/>
          <xsl:choose>
            <xsl:when test="$propertyName = 'Condition'" >
              <xsl:if test="$propertyValue != ''">
                <xsl:value-of select="$newline"/>
                <xsl:element name="Condition" namespace="http://schemas.microsoft.com/wix/2006/wi">
                  <xsl:value-of select="$propertyValue"/>
                </xsl:element>
              </xsl:if>
            </xsl:when>
            <xsl:when test="$propertyName = 'HKCUKey'" >
              <xsl:if test="$propertyValue != ''">
                <xsl:value-of select="$newline"/>
                <xsl:element name="RegistryValue" namespace="http://schemas.microsoft.com/wix/2006/wi">
                  <xsl:attribute name="Root">
                    <xsl:text>HKCU</xsl:text>
                  </xsl:attribute>
                  <xsl:attribute name="Key">
                    <xsl:value-of select="$propertyValue"/>
                  </xsl:attribute>
                  <xsl:attribute name="Name">
                    <xsl:value-of select="$componentId"/>
                  </xsl:attribute>
                  <xsl:attribute name="Type">
                    <xsl:text>string</xsl:text>
                  </xsl:attribute>
                  <xsl:attribute name="Value">
                    <xsl:call-template name="filename-with-extension">
                      <xsl:with-param name="path" select="$fileSource"/>
                    </xsl:call-template>
                    <xsl:text> file</xsl:text>
                  </xsl:attribute>
                  <xsl:attribute name="KeyPath">
                    <xsl:text>yes</xsl:text>
                  </xsl:attribute>
                </xsl:element>
              </xsl:if>
            </xsl:when>
            <xsl:when test="$propertyName = 'FileTypes'" >
              <xsl:variable name="nodeSet" select="$comFileNode//*[name()=$propertyName]" />
              <xsl:if test="count($nodeSet//wix:*) > 0">
                <xsl:value-of select="$newline"/>
                <xsl:for-each select="$nodeSet//wix:*[parent::*[name()='FileTypes']]" >
                  <xsl:copy-of select="."/>
                </xsl:for-each>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="$propertyValue != ''">
                <xsl:attribute name="{$propertyName}">
                  <xsl:value-of select="$propertyValue"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
        <!-- Search for the 'TypeLib' element with the parent equals to the current 'Component' element -->
        <xsl:variable name="typeLibNode" select=".//wix:TypeLib[parent::*[name()=name(current())]]" />
        <xsl:for-each select="*">
          <xsl:choose>
            <!-- Copy the 'File' element -->
            <xsl:when test="name(.)='File'">
              <xsl:value-of select="$newline"/>
              <!-- Create the 'File' element -->
              <xsl:element name="{name(.)}" namespace="http://schemas.microsoft.com/wix/2006/wi">
                <xsl:choose>
                  <xsl:when test="$fileSource = ''">
                    <!-- Copy all attributes -->
                    <xsl:call-template name="copy-all-attributes">
                      <xsl:with-param name="selection" select="./@*"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:choose>
                      <xsl:when test="$shareFileId != ''">
                        <xsl:attribute name="Id">
                          <xsl:value-of select="$shareFileId"/>
                        </xsl:attribute>
                        <!-- Copy all attributes except for 'Id' and 'Source' (we create it manually below) -->
                        <xsl:call-template name="copy-all-attributes">
                          <xsl:with-param name="selection" select="./@*[(name() != 'Id') and (name() != 'Source')]"/>
                        </xsl:call-template>
                      </xsl:when>
                      <xsl:otherwise>
                        <!-- Copy all attributes except for 'Source' (we create it manually below) -->
                        <xsl:call-template name="copy-all-attributes">
                          <xsl:with-param name="selection" select="./@*[name() != 'Source']"/>
                        </xsl:call-template>
                      </xsl:otherwise>
                    </xsl:choose>
                    <xsl:attribute name="Source">
                      <xsl:value-of select="$fileSource"/>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!-- Create attributes from the 'file-attributes' array -->
                <!-- All values are loaded from the 'COMEntry' node located in the 'RegisterForCOM.xml' file -->
                <xsl:for-each select="$file-attributes">
                  <xsl:variable name="propertyName" select="."/>
                  <xsl:variable name="propertyValue" select="normalize-space($comFileNode//*[name()=$propertyName])"/>
                  <xsl:if test="$propertyValue!=''">
                    <xsl:attribute name="{$propertyName}">
                      <xsl:value-of select="$propertyValue"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:for-each>
                <!-- We can copy all elements at once: <xsl:copy-of select="*|@*"/> -->
                <!-- But we need to create 'TypeLib' manually -->
                <xsl:for-each select="*">
                  <xsl:value-of select="$newline"/>
                  <xsl:choose>
                    <xsl:when test="name(.)='TypeLib'">
                      <xsl:call-template name="create-typelib">
                        <xsl:with-param name="typeLibOriginal" select="."/>
                        <xsl:with-param name="directoryId" select="$directoryId"/>
                        <xsl:with-param name="fileId" select="$shareFileId"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="$shareFileId != ''">
                          <xsl:call-template name="copy-and-replace">
                            <xsl:with-param name="sourceElem" select="."/>
                            <xsl:with-param name="fileId" select="$shareFileId"/>
                          </xsl:call-template>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:copy-of select="."/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
                <!-- If 'TypeLib' with parent equals to 'Component' was found, move it to the 'File' element -->
                <xsl:if test="$typeLibNode">
                  <xsl:value-of select="$newline"/>
                  <xsl:call-template name="create-typelib">
                    <xsl:with-param name="typeLibOriginal" select="..//wix:TypeLib"/>
                    <xsl:with-param name="directoryId" select="$directoryId"/>
                    <xsl:with-param name="fileId" select="$shareFileId"/>
                  </xsl:call-template>
                </xsl:if>
                <xsl:value-of select="$newline"/>
              </xsl:element>
            </xsl:when>
            <!-- Copy all nodes except for 'TypeLib' ('TypeLib' should be moved to the correct location by the code above) -->
            <xsl:when test="name(.)!='TypeLib'">
              <xsl:value-of select="$newline"/>
              <xsl:choose>
                <xsl:when test="$shareFileId != ''">
                  <!-- Create a new element and replace file references with the correct value -->
                  <xsl:call-template name="copy-and-replace">
                    <xsl:with-param name="sourceElem" select="."/>
                    <xsl:with-param name="fileId" select="$shareFileId"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Create a copy of the current element -->
                  <xsl:element name="{name(.)}" namespace="http://schemas.microsoft.com/wix/2006/wi">
                    <xsl:call-template name="copy-all-attributes">
                      <xsl:with-param name="selection" select="@*"/>
                    </xsl:call-template>
                  </xsl:element>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
        <xsl:value-of select="$newline"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>