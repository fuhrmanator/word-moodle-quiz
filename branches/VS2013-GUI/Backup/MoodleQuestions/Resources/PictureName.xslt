<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
	xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint">

	<xsl:output method="text" /> 

	<!-- remove all text -->
  <xsl:template match="text()"></xsl:template> 


  <xsl:template match="w:p/w:r/w:pict"><xsl:value-of select="w:binData/@w:name"/></xsl:template>

</xsl:stylesheet>