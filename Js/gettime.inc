<script language="vbscript" RUNAT="Server">
 function getAtualTime
  getAtualTime = Now()
 End function 

</script>
<script language="jscript" RUNAT="Server">
 
 function GetTimeTest(testdate)
{
  var d,  t;
  var MinMilli = 1000 * 60;
  var HrMilli = MinMilli * 60;
  var DyMilli = HrMilli * 24;
  d = new Date();
  t = Date.parse(testdate);
  
  return(Math.round(Math.abs(t / DyMilli)));
}


 function CompareDate(date1, date2)
{
  return(Math.abs((GetTimeTest(date2) - GetTimeTest(date1))));
}
 
 function Time()
{
  var d, s = "";
  var c = ":";
  d = new Date();
  
  s += d.getHours() + c;
  s += d.getMinutes() + c;
  s += d.getSeconds() + c;
  s += d.getMilliseconds();
  return(s);
}

function Today()
{
  var d, s = "";
  
  d = new Date();
    
  s += (d.getMonth() + 1) + "/";
  s += d.getDate() + "/";
  s += d.getYear();
  return(s);
}

function CurrentYear()
{
  var d, s = "";
  
  d = new Date();
    
  s += d.getYear();
  return(s);
}

function NowForQuery()
{


  var d, s = Application("dchar");
  var c = ":";
  
  d = new Date();
    
  s += (d.getMonth() + 1) + "/";
  s += d.getDate() + "/";
  s += d.getYear();
  s += " ";
  s += d.getHours() + c;
  s += d.getMinutes() + c;
  s += d.getSeconds();
  s += Application("dchar"); 
  return(s);
}

</script>
