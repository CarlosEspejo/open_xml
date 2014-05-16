<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                              xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

 <xsl:template match="p">
   <w:p>
     <w:r>
      <w:t><xsl:value-of select="."/></w:t>
    </w:r>
   </w:p>
  </xsl:template>
</xsl:stylesheet>
