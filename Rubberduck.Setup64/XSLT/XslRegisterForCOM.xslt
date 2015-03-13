<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" >
  <!-- The name of the operation that is used in the 'COMEntry' template -->
  <xsl:param name="operationType" select="''"></xsl:param>
  <!-- The intermediate directory of the setup project -->
  <xsl:param name="intermediateDir" select="''"></xsl:param>

  <xsl:output method='text'/>
  <xsl:strip-space elements='*'/>

  <xsl:variable name='newline'>
    <xsl:text>&#13;&#10;</xsl:text>
  </xsl:variable>

  <xsl:variable name='file-extension-delimiter'>
    <xsl:text><![CDATA[.]]></xsl:text>
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
  <xsl:template name="extract-fileName-withoutextension">
    <xsl:param name="path" />
    <xsl:choose>
      <xsl:when test="contains($path,'\')">
        <xsl:call-template name="extract-fileName-withoutextension">
          <xsl:with-param name="path" select="substring-after($path,'\')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($path,'/')">
        <xsl:call-template name="extract-fileName-withoutextension">
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

  <!-- Create a text file based on the information from the 'COMEntry' nodes of 'RegisterForCOM.xml' file -->
  <xsl:template match="COMEntry" >
    <!-- Create a text file that contains the list of COM files -->
    <!-- The list is used to extract the registry information from COM files using the Heat tool -->
    <xsl:if test="$operationType = 'HeatFiles'">
      <xsl:value-of select="concat(@Source, $newline)" disable-output-escaping="yes"/>
    </xsl:if>
    <!-- Create a text file that contains the list of '*.com.xml' files -->
    <xsl:if test="$operationType = 'TransformFiles'">
      <xsl:variable name="sourceFileName">
        <xsl:call-template name="extract-fileName-withoutextension">
          <xsl:with-param name="path" select="@Source"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:value-of select="concat($intermediateDir, '_', $sourceFileName, '.com.xml', $newline)" disable-output-escaping="yes"/>
    </xsl:if>
    <!-- Create a text file with the list of '*.com.wsx' files which should be deleted -->
    <xsl:if test="$operationType = 'DeleteFiles'">
      <xsl:if test="@Type='ProjectOutput'">
        <xsl:variable name="sourceFileName">
          <xsl:call-template name="extract-fileName-withoutextension">
            <xsl:with-param name="path" select="@Source"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat($intermediateDir, '_', $sourceFileName, '.com.wsx', $newline)" disable-output-escaping="yes"/>
      </xsl:if>
    </xsl:if>
    <!-- Create a text file with the list of '*.com.wsx' files which should be included in the setup project for compilation -->
    <xsl:if test="$operationType = 'CompileFiles'">
      <xsl:if test="@Type='File'">
        <xsl:variable name="sourceFileName">
          <xsl:call-template name="extract-fileName-withoutextension">
            <xsl:with-param name="path" select="@Source"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat($intermediateDir, '_', $sourceFileName, '.com.wsx', $newline)" disable-output-escaping="yes"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>