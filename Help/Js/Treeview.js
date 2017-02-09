function GetChildElem( eSrc,sTagName ) {
    	var listArr = eSrc.children;
	for ( var i=0; i < listArr.length; i++ ) {
		if ( sTagName == listArr[i].tagName ) return listArr[i];
    	}
    	return false;
}

function Do_Expand_All() {
	var listArr = document.body.getElementsByTagName( "LI" );
  	for ( var i=0; i < listArr.length; i++ ) {
	  	if ( listArr[i].className == "clsHasKidsCollapsed" ) {	  	
			var eChild = GetChildElem( listArr[i], "UL" );
			var aChild = GetChildElem( listArr[i], "A" );		        
		        listArr[i].className = "clsHasKidsExpanded";
		        eChild.style.display = "block";  	
	  	}
  	}
}

function Do_Collapse_All() {
	var listArr = document.body.getElementsByTagName( "LI" );
  	for ( var i=0; i < listArr.length; i++ ) {
	  	if ( listArr[i].className == "clsHasKidsExpanded" ) {	  	
			var eChild = GetChildElem( listArr[i], "UL" );
			var aChild = GetChildElem( listArr[i], "A" );		        
		        listArr[i].className = "clsHasKidsCollapsed";
		        eChild.style.display = "none";  	
	  	}
  	}
}

function FindObjectByName( objName, objTagName ) {
    	var listArr = document.body.getElementsByTagName( objTagName );
	for ( var i=0; i < listArr.length; i++ ) {
		if ( listArr[i].name == objName ) return listArr[i];
    	}
    	return false;
}

function ListContainSearchedNode( NodeMame ) {
	var listArr = document.body.getElementsByTagName( "LI" );
  	for ( var i=0; i < listArr.length; i++ ) {
  		if ( GetChildElem( listArr[i], "A" ).name ==  NodeMame ) {
  			return GetChildElem( listArr[i], "A" ).parentElement;
		}
  	}
}

function GetLink( NodeMame ) {
	var listArr = document.body.getElementsByTagName( "A" );
  	for ( var i=0; i < listArr.length; i++ ) {
  		if ( listArr[i].name ==  NodeMame ) {
  			return listArr[i];
		}
  	}
}
function LocateChild( childValue ) { 
	var SearchedLesson = "Lesson" + childValue; 
 	var ObjSrc = ListContainSearchedNode( SearchedLesson );
 	if ( ObjSrc ) {
 		var aChild = GetChildElem( ObjSrc, "A" );
 		//aChild.className = "ActiveLessonLnk";
 		while ( ObjSrc.parentElement.parentElement.tagName != "BODY" ) {
			ObjSrc = ObjSrc.parentElement.parentElement;
			ObjSrc.className = "clsHasKidsExpanded";
			var eChild = GetChildElem( ObjSrc, "UL" );
			eChild.style.display = "block";
		}
		location = "#" + SearchedLesson;		
		//aChild.setActive();
		aChild.focus();
	}		
}

function RestoreLastLink( childValue ) {  
	var SearchedLesson = "Lesson" + childValue; 
 	var ObjSrc = ListContainSearchedNode( SearchedLesson );
 	if ( ObjSrc ) {
 		var aChild = GetChildElem( ObjSrc, "A" );
 		//aChild.className = "LessonLnk";
	}	
}