var checked_all = false;

function GetObjectCollection( pContainer, pObjectType ) {
	if ( pContainer ) {
		return document.body.getElementsByTagName( pObjectType );
	} else {
		return pContainer.getElementsByTagName( pObjectType );
	}	
}

function CheckAll( pContainer ) {
	var CheckList = GetObjectCollection( pContainer, "INPUT" );
  	for ( var i=0; i < CheckList.length; i++ ) {
	  	if ( CheckList[i].type == "checkbox" ) {	  	
	  		CheckList[i].checked = !checked_all;
	  	}
  	}
  checked_all = !checked_all;
}

function UnCheckAll( pContainer ) {
	var CheckList = GetObjectCollection( pContainer, "INPUT" );
  	for ( var i=0; i < CheckList.length; i++ ) {
	  	if ( CheckList[i].type == "checkbox" ) {	  	
	  		CheckList[i].checked = false;
	  	}
  	}
}