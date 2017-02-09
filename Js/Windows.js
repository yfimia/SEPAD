function abreVentana( pName, pwidth, pheight, pUrl, pScroll, pResize) {   	
	var cad = "width=" + pwidth + ",height=" + pheight + ",toolbar=no,menubar=no,location=no,scrollbars=";
     	if (pScroll) {
		cad = cad + pScroll +",resizable=yes,status=no";
     	} else {
     		cad = cad + "yes,";
     	} 	
     	if (pResize) {
     		cad = cad + ",resizable=" + pResize + ",status=no";
     	} else {
    		cad = cad + "resizable=yes,status=no";
     	} 	
   	vent = open(pUrl, pName, cad);     
}