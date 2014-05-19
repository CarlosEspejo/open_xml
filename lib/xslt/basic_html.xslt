<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                              xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <xsl:output method="xml"/>

  <xsl:template match="p">
    <w:p>
      <xsl:apply-templates />
    </w:p>
  </xsl:template>

  <xsl:template match="text()">
    <xsl:if test="string-length(.) > 0">
      <xsl:choose>
       <xsl:when test="parent::strong or parent::b">
          <w:r>
            <w:rPr>
              <w:b />
            </w:rPr>
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </w:r>
        </xsl:when>
         <xsl:when test="parent::em or parent::i">
          <w:r>
            <w:rPr>
              <w:i />
            </w:rPr>
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </w:r>
        </xsl:when>
        <xsl:otherwise>
          <w:r>
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </w:r>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
