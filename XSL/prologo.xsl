<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/TR/WD-xsl">

  <xsl:template match="/">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>
            
            <xsl:apply-templates select="Prologo/PalabrasClaves" />
            <br/>            
            
            <xsl:apply-templates select="Prologo/Fundament" />
            <br/>

            <xsl:apply-templates select="Prologo/Calendario" />            
            <br/>
            
            <xsl:apply-templates select="Prologo/Objetivos" />
            <br/>

            <xsl:apply-templates select="Prologo/Programa" />
            <br/>

            

            <xsl:apply-templates select="Prologo/NotasAdicionales" />            
            <br/>

            <xsl:apply-templates select="Prologo/SistemaTutorias" />            
            <br/>
            
            <xsl:apply-templates select="Prologo/ConocimientosRequeridos" />            
            <br/>
            
            <xsl:apply-templates select="Prologo/RequerimientosTecnicos" />
            <br/>
          
            <xsl:apply-templates select="Prologo/OrientacionesGenerales" />
            <br/>

            <xsl:apply-templates select="Prologo/Bibliog" />
            <br/>
            <xsl:apply-templates select="Prologo/Recursos" />
            <br/>

          </td>
        </tr>
      </table>
      
  </xsl:template>

  <xsl:template match="Titulo">
    <tr> 
      <td> 
         <h2> <xsl:apply-templates /> </h2>
      </td>
    </tr>

  </xsl:template>
  <xsl:template match="Autor">
    <tr> 
       <td><b>Autor: </b><xsl:apply-templates /> </td>
    </tr>
  </xsl:template>

  <xsl:template match="PalabrasClaves[node()]">
      
      <table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
        <tr> 
          <td class="ToolBar"><b>Palabras clave:</b> </td>
        </tr>
        <tr> 
         <td>
           <xsl:apply-templates >
                 <xsl:template match="PalabraClave">
                     
                       | <xsl:apply-templates /> |
                     
                  </xsl:template>    
           </xsl:apply-templates >                         
           </td>
        </tr>
      </table>
      
  </xsl:template>

  <xsl:template match="Objetivos[node()]">
        <table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
           <tr> 
             <td class="ToolBar"><b>Objetivos Generales:</b> </td>
           </tr>
           <tr> 
              <td>
        		  <ul> 
                   <xsl:apply-templates >
                     <xsl:template match="Objetivo">
                       <li>
                         <xsl:apply-templates />
                       </li>
                     </xsl:template>
                   </xsl:apply-templates >
				  </ul> 					
              </td>
           </tr>
        </table>
  </xsl:template>

  <xsl:template match="ConocimientosRequeridos[node()]">
       <table width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> 
               <b>Conocimientos Requeridos:</b> 
            </td>
          </tr>
           <tr> 
             <td> 
      		   <ul> 				
                 <xsl:apply-templates >
                    <xsl:template match="Conocimiento">
                       <li>
                         <xsl:apply-templates />
                        </li>
                     </xsl:template>
                 </xsl:apply-templates >      		   
 			   </ul> 					
             </td>
           </tr>
       </table>

  </xsl:template>


  <xsl:template match="Calendario[node()]">
            <table border="0" cellpadding="3" cellspacing="1" width="90%" align="center" class="PrologueTable">
              <tr> 
                <td class="ToolBar" colspan="2"> <b>Calendario de Actividades 
                  </b> </td>
              </tr>
              <xsl:apply-templates >
               <xsl:template match="Actividad">
                 <tr>
                    <xsl:apply-templates>
                        <xsl:template match="text()">
                          <td width="80%" height="16"><xsl:value-of /></td>
                        </xsl:template>
                    </xsl:apply-templates>
                    <xsl:apply-templates select="@fecha" />
                 </tr>
             </xsl:template>
             <xsl:template match="@fecha">
               <td width="20%"><xsl:apply-templates/></td>
             </xsl:template>
        </xsl:apply-templates >
              
     </table>

  </xsl:template>

  <xsl:template match="RequerimientosTecnicos[node()]">
        <table border="0" cellpadding="3" cellspacing="1" align="center" width="90%" class="PrologueTable">
  	    <tr> 
	      <td class="ToolBar" colspan="2"> <b>Requerimientos técnicos y de software </b> </td>
            </tr>
        <xsl:apply-templates >
          <xsl:template match="Requerimiento">
             <tr>
                <xsl:apply-templates>
                    <xsl:template match="text()">
                      <td width="55%"><xsl:value-of /></td>
                    </xsl:template>
                </xsl:apply-templates>
                <xsl:apply-templates select="@url" />
             </tr>
          </xsl:template>
          <xsl:template match="@url">
            <td width="45%"><xsl:apply-templates/></td>
          </xsl:template>

        </xsl:apply-templates >
       </table>     
  </xsl:template>

  <xsl:template match="OrientacionesGenerales[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Orientaciones Generales </b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>
  </xsl:template>

  <xsl:template match="NotasAdicionales[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Sistema de Evaluación</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>

  <xsl:template match="Fundament[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Fundamentación</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>


  <xsl:template match="Recursos[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Recursos</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>

  <xsl:template match="Programa[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Programa</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>
  
  <xsl:template match="Bibliog[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Bibliografía General</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>  
  
  <xsl:template match="SistemaTutorias[text()]">
       <table border="0" width="90%" cellpadding="2" cellspacing="1" align="center" class="PrologueTable">
          <tr> 
            <td class="ToolBar"> <b>Sistema de Tutorías</b> </td>
          </tr>
          <tr> 
             <td> 
               <p align="justified"><xsl:apply-templates /></p>
             </td>
           </tr>
      </table>  
  </xsl:template>    
  <xsl:template  match="BR"><BR/></xsl:template>
  <xsl:template  match="BR"><BR/></xsl:template>

  <xsl:template  match="text()"><xsl:value-of  /></xsl:template>

</xsl:stylesheet>