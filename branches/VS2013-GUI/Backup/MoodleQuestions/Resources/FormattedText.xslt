<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
	xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint">

	<xsl:output method="text" /> 

	<!-- remove all text -->
  <xsl:template match="text()"></xsl:template> 



  <xsl:template match="w:p/w:r//node() [@w:val='preformatted']">&lt;pre&gt;<xsl:value-of select="../../w:t"/>&lt;/pre&gt;</xsl:template>

  <xsl:template match="w:p/w:r//node() [@w:val='MissingWord']">__________</xsl:template>

  <xsl:template match="w:p/w:r [not(w:rPr/w:rStyle)]"><xsl:value-of select="w:t"/></xsl:template>

</xsl:stylesheet>
