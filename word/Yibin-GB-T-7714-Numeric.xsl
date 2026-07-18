<?xml version="1.0" encoding="utf-8"?>
<!--
  YibinThesis Word bibliography compatibility style.

  Word ships a GB7714 bibliography formatter as GB.XSL.  This small wrapper
  keeps that formatter for bibliography entries and replaces only the inline
  citation branch with a numeric [n] form.  It contains no Microsoft source
  code; GB.XSL remains an external prerequisite supplied by Microsoft Word.
-->
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
  <xsl:import href="GB.XSL"/>
  <xsl:output method="html" encoding="utf-8"/>

  <xsl:template match="/">
    <xsl:choose>
      <xsl:when test="b:StyleNameLocalized">
        <xsl:text>GB/T 7714 数字引用（YibinThesis）</xsl:text>
      </xsl:when>
      <xsl:when test="b:Citation">
        <html xmlns="http://www.w3.org/TR/REC-html40">
          <body>
            <p>
              <xsl:for-each select="b:Citation">
                <xsl:if test="b:FirstAuthor"><xsl:text>[</xsl:text></xsl:if>
                <xsl:choose>
                  <xsl:when test="b:Source/b:RefOrder">
                    <xsl:value-of select="b:Source/b:RefOrder"/>
                  </xsl:when>
                  <xsl:otherwise><xsl:value-of select="b:Source/b:Tag"/></xsl:otherwise>
                </xsl:choose>
                <xsl:choose>
                  <xsl:when test="b:LastAuthor"><xsl:text>]</xsl:text></xsl:when>
                  <xsl:otherwise><xsl:text>,</xsl:text></xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </p>
          </body>
        </html>
      </xsl:when>
      <xsl:otherwise><xsl:apply-imports/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
