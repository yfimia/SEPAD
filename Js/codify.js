  key = 10;
    
  function code( st )
   {
    var stcode = "";char = ""; 
    for (i=0;i < st.length; i++)
     {
      char   = String.fromCharCode( (st.charCodeAt(i) + key) % 256 );
      stcode = stcode + char; 
     } 
    return stcode
   }
   
  function decode( st )
   {
    var stcode = "";char = ""; 
    for (i=0;i < st.length; i++)
     {
      char   = String.fromCharCode( (255*key + st.charCodeAt(i)) % 256);
      stcode = stcode + char; 
     } 
   return stcode
  }