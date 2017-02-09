<%
 function change(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "\n";
    	    brtext = "<BR/>";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext); 
        }
 return(s);        
}

function unchange(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "<br/>";
    	    brtext = "\n";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
    	var cl = "<BR/>";
    	    brtext = "\n";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 

 return(s);        
}

function unchange1(ss) {
    	var cnt = 0;
    	s = new String(ss);
    	var cl = "<br/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 
    	var cl = "<BR/>";
    	    brtext = "";
    	while (s.search(cl) != -1) {
         	s = s.replace(cl,brtext);
        } 

 return(s);        
}

%>