<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">

<xsl:template match="/">
	<HTML>
	<HEAD> 
	<TITLE>test</TITLE>
	<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="idx.js"></SCRIPT>
	
	</HEAD>
	<BODY>
	<xsl:value-of select="idx/k" />
	<xsl:apply-templates select="idx/p" />
	</BODY>
	</HTML>
</xsl:template>

<xsl:template match="p">

<input type="button" >
<xsl:attribute name="value"><xsl:value-of select="t" /></xsl:attribute>
<xsl:attribute name="onclick">toggle(<xsl:value-of select="t" />)</xsl:attribute>
</input>
<span>
<xsl:attribute name="id"><xsl:value-of select="t" /></xsl:attribute>

<xsl:for-each select="s">
	<span style="background-color: #FFFF00">*</span>
	<xsl:for-each select="w">
		<span style="background-color: #E0E0E0">
		<xsl:value-of select="." /></span>
	</xsl:for-each>
	<xsl:value-of select="g" />

</xsl:for-each>

</span>
<br/>

</xsl:template>
</xsl:stylesheet>