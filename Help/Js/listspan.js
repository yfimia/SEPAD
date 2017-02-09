  function GetChildElem(eSrc,sTagName)
  {
    var cKids = eSrc.children;
    for (var i=0;i<cKids.length;i++)
    {
      if (sTagName == cKids[i].tagName) return cKids[i];
    }
    return false;
  }
  
  function document.onclick()
  { 
    var eSrc = window.event.srcElement;
		if (eSrc.tagName == "LI" && "clsHasKidsCollapsed" == eSrc.className || "clsHasKidsExpanded" == eSrc.className)
		{  
			var eChild = GetChildElem(eSrc,"UL");
			var aChild = GetChildElem(eSrc,"A");
		        eSrc.className = (("block" == eChild.style.display) ? "clsHasKidsCollapsed" : "clsHasKidsExpanded");
		        eChild.style.display = ("block" == eChild.style.display ? "none" : "block");
		}
  }
