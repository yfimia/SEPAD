
    //Selection
    function EType1(Options, Texto, ID, Color1, Color2, Color3, Color4)
     // Color1:Color de la tabla exterior
     // Color2:Color de las celdas
     // Color3:Color de la pregunta
     // Color3:Color de la pregunta
      {
        resp = response.split('|||');
        document.write('<table width="100%" border="0" cellspacing="0" cellpadding="0">');
        document.write('  <tr><td>' + Texto + '</td></tr>');
        document.write('  <tr class="ColumnConnectEx">'); 
        document.write('    <td width="50%" align="right">'); 
        document.write('      <table width="98%" border="0" cellspacing="0" cellpadding="2">');
        for (var i = 0;i < Options.length;i++)
          {
        document.write('        <tr>'); 
        document.write('           <td width="1%">'); 
        if ((i < resp.length) && (resp[i] == "true"))
          document.write('             <input type="checkbox" name="' + 'O' + ID + '#' + i + '" value="checkbox" checked >');
        else
          document.write('             <input type="checkbox" name="' + 'O' + ID + '#' + i + '" value="checkbox" >');
        document.write('           </td>');
        document.write('           <td width="99%">' + Options[i] + '</td>');
        document.write('        </tr>');
          }      
        document.write('      </table>');
        document.write('      <hr noshade>');
        document.write('    </td>');
        document.write('  </tr>');
        document.write('</table>');
        document.write('<br>');
      }

    //Enlazar columnas
    function EType2(Option1, Option2, Texto, ID, Color1, Color2, Color3, Color4)
     // Color1:Color de la tabla exterior
     // Color2:Color de las celdas
     // Color3:Color de la pregunta
     // Color4:Color del texto de las celdas
      { 
        resp = response.split('|||');

        var max = 0;
        if (Option1.length > Option2.length)
          max = Option1.length
        else
          max = Option2.length;  
      
        document.write('<table width="100%" border="0" cellspacing="0" cellpadding="0">');
        document.write('  <tr> ');
        document.write('    <td>' + Texto + '</td>');
        document.write('  </tr>');
        document.write('  <tr>'); 
        document.write('    <td align="right">'); 
        document.write('      <table width="98%" border="0" cellspacing="0" cellpadding="2">');
        document.write('        <tr> ');
        document.write('          <td> ');
        document.write('            <table border="0" cellspacing="1" cellpadding="0">');
		document.write('              <tr> ');
        if (Option1.length != 0)
          {
		document.write('                <td width="50%"> ');
		document.write('                  <table width="100%" border="0" cellspacing="0" cellpadding="0">');
		document.write('                    <tr> ');
		document.write('                      <td colspan="2"> <b>Columna &nbsp;&nbsp;A</b> </td>');
		document.write('                    </tr>');
		document.write('                  </table>');
		document.write('                </td>');
		document.write('                <td width="50%"> ');
          } 
        else
		document.write('                <td width="100%"> ');

		document.write('                  <table width="100%" border="0" cellspacing="0" cellpadding="0">');
		document.write('                    <tr> ');
		document.write('                      <td colspan="2"> <b>Columna &nbsp;B</b> </td>');
		document.write('                    </tr>');
		document.write('                  </table>');
		document.write('                </td>');
		document.write('              </tr>');
		
        for (var i = 0;i < max;i++)
           {
            if ((i < Option1.length) && (i < Option2.length))
              { 
               opt1 = Option1[i];
               opt2 = Option2[i];
              }
            else
              if (i >= Option1.length)
                {
                 opt1 = "&nbsp;";
                 opt2 = Option2[i];
                } 
              else  
               if (i >= Option2.length)
                 {
                   opt1 = Option1[i];
                   opt2 = "&nbsp;";
                   
                 }  

    		document.write('              <tr> ');
            if (Option1.length != 0)
              {

		document.write('                <td width="50%"> ');
		document.write('                  <table width="100%" border="0" cellspacing="0" cellpadding="2">');
		document.write('                    <tr> ');
		document.write('                      <td width="1%"> ');
		if (i < Option1.length)
                  {
		document.write('                        <div align="center">' + (i + 1) + '.-</div>');
                  }
                else
                  {
		document.write('                        <div align="center"></div>');
                  }

		document.write('                      </td>');
		document.write('                      <td width="99%">' + opt1 + '</td>');
		document.write('                    </tr>');
		document.write('                  </table>');
		document.write('                </td>');
		document.write('                <td width="50%"> ');
              }
             else
                document.write('                <td width="100%"> ');
		document.write('                  <table width="100%" border="0" cellspacing="0" cellpadding="2">');
		document.write('                    <tr> ');

		if (i < Option2.length)
          {
		document.write('                      <td width="1%"> ');
		document.write('                        <select name="O' + ID + '#' + i + '">');
		document.write('                          <option value="-1" selected >-</option>');
            
		    for (var j = 1;j <= Option1.length;j++)
               {
                 if ((i < resp.length) && (j == resp[i]))
		document.write('                          <option selected value="' + j + '">' + j + '</option>');
		         else
		document.write('                          <option value="' + j + '">' + j + '</option>');
		         
		       }
		        
		document.write('                        </select>');
		document.write('                      </td>');
		  }
		  
		document.write('                      <td width="99%">' + opt2 + '</td>');
		document.write('                    </tr>');
		document.write('                  </table>');
		document.write('                </td>');
		document.write('              </tr>');
		   }
		   
		document.write('            </table>');
		document.write('          </td>');
		document.write('        </tr>');
		document.write('      </table>');
		document.write('      <hr noshade>');
		document.write('    </td>');
		document.write('</tr>');
		document.write('</table>');
		document.write('<br>      ');
      }
      
    function ExistBreak(b,p1,p2)
      {
        var posi = -1;
        for (var i = 0;i < b.length;i++)
          {
            if ((b[i] >= p1) && (b[i] <= p2)) 
              {
                if (posi == -1) 
                  {
                    posi = i; 
                  }
                else
                  if (b[i] < b[posi]) 
                    posi = i;
              }  
          }
        return posi;  
      }
    
 //Completar espacios
 function EType3(Pregunta,ID, Texto, Positions, Lengths, Breaks, Color1,Color2,Color3)  
     // Color1:Color de la tabla exterior
     // Color2:Color del texto
     // Color3:Color del encabezamiento
      {
        resp = response.split('|||');

        var p = 0;
        var last = 0;
        var pp = 0;


		document.write('<table width="100%" border="0" cellspacing="0" cellpadding="0">');
		document.write('  <tr> ');
		document.write('    <td>' + Pregunta + '</td>');
		document.write('  </tr>');
		document.write('  <tr> ');
        document.write('    <td align="right"> ');
		document.write('      <table width="98%" border="0" cellspacing="5" cellpadding="0" align="center">');
        document.write('        <tr>');
        document.write('          <td>');


        while (p < Positions.length)
          {
            pp = ExistBreak(Breaks,last,Positions[p]); 
            while (pp != -1)
              {
                document.write(Texto.substr(last,Breaks[pp] - last + 1) + "<br>");
                last = Breaks[pp] + 1;
                pp = ExistBreak(Breaks,last,Positions[p]); 
              }
            if (last <= Positions[p])
              {
                document.write(Texto.substr(last,Positions[p] - last + 1));
                last = Positions[p] + 1;
              }  
            document.write("<input SIZE='" + Lengths[p] + "' TYPE='TEXT' NAME='O" + ID + "#" + p + "' value='" + resp[p] + "'  class='Edit'>");
            
            p++;
          }
        //alert("final : " + Texto.substr(last,Texto.length - last + 1));  

          pp = ExistBreak(Breaks,last,Texto.length); 
          while (pp != -1)
            {
              document.write(Texto.substr(last,Breaks[pp] - last + 1) + "<br>");
              last = Breaks[pp] + 1;
              pp = ExistBreak(Breaks,last,Texto.length); 
            }

        document.write(Texto.substr(last,Texto.length - last + 1));
         
        document.write('          </td>');
        document.write('        </tr>');
        document.write('      </table>');
        document.write('      <hr noshade>');
        document.write('    </td>');
        document.write('  </tr>');
        document.write('</table>');
        document.write('<br>');
      }

    //Ordenar no esta terminado
    function EType4(Option1, Texto, ID, Color1, Color2, Color3, Color4)
     // Color1:Color de la tabla exterior
     // Color2:Color de las celdas
     // Color3:Color de la pregunta
     // Color4:Color del texto de las celdas
      {
        document.write("<H1>" + hola + Texto + "</H1>");
        document.write("<table border=2 bordercolor=" + Color1 +">");
        for (var i = 0;i < Option1.length;i++)
          {
            document.write("<tr bordercolor=" + Color2 + ">");         
            document.write("<td><font color=" + Color4 + ">" + "<INPUT SIZE=1 TYPE=TEXT NAME='O" + ID + "#" + i + "'>" +"</font></td>");
            document.write("<td><font color=" + Color4 + ">" + Option1[i] + "</font></td>");
            document.write("</tr>");
          }
        document.write("</table>");
        document.write("<br><hr>");
      }
    
    //No me queda claro
    function EType5(ID, Texto, Positions, Lengths, Breaks, Color1,Color2)  
     // Color1:Color de la tabla exterior
     // Color2:Color del texto
      {
        document.write("<table border=2 bordercolor=" + Color1 + "><td>");
        document.write("<P STYLE='COLOR:" + Color2 + "'>");
        var p = 0;
        var last = 0;
        var pp = 0;
        while (p < Positions.length)
          {
            pp = ExistBreak(Breaks,last,Positions[p]); 
            while (pp != -1)
              {
                document.write(Texto.substr(last,Breaks[pp] - last + 1) + "<br>");
                last = Breaks[pp] + 1;
                pp = ExistBreak(Breaks,last,Positions[p]); 
              }
            if (last <= Positions[p])
              {
                document.write(Texto.substr(last,Positions[p] - last + 1));
                last = Positions[p] + 1;
              }  
            document.write("<INPUT SIZE=" + Lengths[p] + " TYPE=TEXT NAME='O" + ID + "#" + p + "'>");
            
            p++;
          }
        //alert("final : " + Texto.substr(last,Texto.length - last + 1));  

          pp = ExistBreak(Breaks,last,Texto.length); 
          while (pp != -1)
            {
              document.write(Texto.substr(last,Breaks[pp] - last + 1) + "<br>");
              last = Breaks[pp] + 1;
              pp = ExistBreak(Breaks,last,Texto.length); 
            }

        document.write(Texto.substr(last,Texto.length - last + 1));
        document.write("</P>");
        document.write("</td></table>");
        document.write("<hr>");
      }

    
    function EType6(ID, Texto,Pwidth,Pheight,Color1,Color2,Color3)
     //Color1: Color de la tabla externa
     //Color2: Color de la pregunta
     //Color1: Color de la respuesta
      {   
		document.write('<table width="100%" border="0" cellspacing="0" cellpadding="0">');
		document.write('  <tr> ');
		document.write('    <td>' + Texto + '</td>');
		document.write('  </tr>');
        document.write('  <tr> ');
		document.write('    <td> ');
		document.write('      <table width="100%" border="0" cellspacing="5" cellpadding="0" align="center">');
		document.write('        <tr> ');
		document.write('          <td>');
		document.write('            <textarea class="TextArea" id="O' + ID + '#0" name="O' + ID + '#0" class="Edit" rows="13">' + response + '</textarea>');
		document.write('          </td>');
		document.write('        </tr>');
		document.write('      </table>');
		document.write('      <hr noshade>');
		document.write('    </td>');
		document.write('  </tr>');
		document.write('</table>');
		document.write('<br>');      
      }  

