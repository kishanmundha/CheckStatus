function getColor(num,max,min) {	// Get Color String

	var range = (max-min)/2

	var range1_min = min;
	var range1_max = min+range;
	
	var range2_min = max-range;
	var range2_max = max;
	
	range = range==0?1:range;
	var colorString = "#";
	
	if(num < range1_max)
	{
		colorString+=convertNum( Math.floor( (num-range1_min)*255/range ) );
		colorString+="FF";
	}
	else
	{
		colorString+="FF";
		colorString+=convertNum( 255-Math.floor( (num-range2_min)*255/range ) );
	}
	
	colorString += "00";
	return colorString;
}

function convertNum(num) {	// Get One Color String Set

	var Hex = new Array('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');
	num = num%256;

	var st = Hex[Math.floor(num/16)] + Hex[num%16];
	
	return st;	
}

function fig2Words1(myNumber) {		// Set figer to Memory
	var x = Number(myNumber);
	if (x/1024 < 1)
	  return x.toFixed(2) + " " + "Bytes";
	else if(x/(1024 * 1024) < 1)
	  return (x/1024).toFixed(2) + " " + "KB";
	else if(x/(1024 * 1024 * 1024) < 1)
	  return (x/(1024 * 1024)).toFixed(2) + " " + "MB";    
	else if(x/(1024 * 1024 * 1024 * 1024) < 1)
	  return (x/(1024 * 1024 * 1024)).toFixed(2) + " " + "GB";    
}

function fig2Words(myNumber) {
	return(fig2Words1(myNumber).replace(".00",""));
}

function ShowUsedSpace(drvPath)
{
	s =0;
	s = ShowTotalSpace(drvPath)-ShowFreeSpace(drvPath);
	return(s);
}

function ShowTotalSpace(drvPath)
{
  var fso, d, s =0;
  fso = new ActiveXObject("Scripting.FileSystemObject");
  d = fso.GetDrive(fso.GetDriveName(drvPath));
  s = d.TotalSize;
  return(s);
}

function ShowFreeSpace(drvPath)
{
  var fso, d, s = 0;
  fso = new ActiveXObject("Scripting.FileSystemObject");
  d = fso.GetDrive(fso.GetDriveName(drvPath));
  s = d.FreeSpace;
  return(s);
}

function init_size()
{

}

function init_drive()
{
	if(loc_path.value) return;
	var fso, s, n, e, x;
	fso = new ActiveXObject("Scripting.FileSystemObject");
	e = new Enumerator(fso.Drives);
	s = "";
	for (; !e.atEnd(); e.moveNext())
	{
		x = e.item();
		if (x.DriveType == 3);
		else if (x.IsReady)
		{
			loc_path.value=x+"\\";
			break;
		}
	}
}

function waitBox_showHide()
{
	switch(arguments[0])
	{
		case 'show' :
			with(waitBox.style)
			{
				pixelLeft	= document.body.offsetWidth/2 - pixelWidth/2;
				pixelTop	= document.body.offsetHeight/2 - pixelHeight/2;
			}
			waitBox.style.visibility = 'visible';
			break;
		
		case 'hide' :
			waitBox.style.visibility = 'hidden';
			break;
		
		default : alert("Error in waitBox showHide Function."); break;
	}
}

function figAttr(a) {
	switch (a) {
		case 6: return "   SH    "; break;
		case 16: return "         "; break;
		case 17: return "     R   "; break;
		case 18: return "    H    "; break;
		case 19: return "    HR  I"; break;
		case 22: return "   SH    "; break;
		case 32: return "A        "; break;
		case 38: return "A  SH   I"; break;
		case 39: return "A  SHR   "; break;
		case 48: return "A        "; break;
		case 1046: return "     R "; break;
		default: return a;
	}
}