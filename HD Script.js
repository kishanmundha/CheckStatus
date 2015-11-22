var tableContent = {
	checkBox: 0,
	itemName: 1,
	itemAttribute: 2,
	itemType: 3,
	itemSize: 4,
	itemStatus: 5
}

var fl_fld = function(name, attr, size, type, isFolder) {	// save properties in this function

	this.attribute = attr;
	this.name = name;
	this.size = size;
	this.type = type;
	this.isFolder = isFolder;
}

var size_max = null;
var size_min = null;
var Total_Size = 0;
var files = new Array();
var fc = null;
var fso = new ActiveXObject("Scripting.FileSystemObject");
var Err = false;

function getNow() {
	var nowTime = new Date();
	var miliS = nowTime.getMilliseconds();
	return Date.parse(nowTime) + miliS;
}

var START_TIME = 0;
function StartTime(T) {
	START_TIME = T;
}

var END_TIME = 0;
function EndTime(T) {
	END_TIME = T;
}

function getUsedTime(T) {
	EndTime(T);
	return (END_TIME - START_TIME);
}

var loadStage = 0;
function load() {	// run First function to load

	switch (loadStage++) {
		case 0:
			if (!loc_path.value) {
				alert("Invalid Path");
				return;
			}
			else if (!checkPath()) return;

			Real_path = loc_path.value;
			waitBox_showHide('show');
			waitMsg.innerText = "Preparing...";
			document.body.scrollTop = 0;
			Total_Size = 0;
			Err = false;

			files = null;
			files = new Array();

			size_max = null;

			size_min = null;
			ShowMemory();
			window.setTimeout("load()", 100);
			break;

		case 1:
			getFolders(1);
			break;

		case 2:
			getFiles(1);
			break;

		case 3:
			shortList(1); 	/* Sorting */
			break;

		case 4:
			ClearRecord(1); 	/* Clear Old list from Table */
			break;

		case 5:
			table1.rows(0).cells(tableContent.checkBox).children(0).checked = false;
			createNew(1); /* Create New List in Table */
			break;

		default:
			loadStage = 0;
			waitBox_showHide('hide');
			return; break;
	}
}

var loadInnerFolder = {
	start: function() {
		if (this.FuncStart) { alert("Inner Folder Size Geting Already Processing!"); return; }
		this.FuncStart = true;
		this.AllSize = 0;
		this.PandingFolder = new Array();
		this.ide = event.srcElement;
		this.TotalSize = Total_Size;

		while (this.ide.tagName != "TR") {
			this.ide = this.ide.parentElement;
		}

		this.PandingFolder[0] = Real_path + this.ide.childNodes.item(tableContent.itemName).innerText + "\\";

		this.getAll();

	},
	getAll: function() {
		if (!this.PandingFolder.length) {
			this.complete();
			return;
		}

		this.getFiles(this.PandingFolder[0]);
		this.getFolders(this.PandingFolder[0]);
		this.showSize();
	},
	getFiles: function(path) {
		try {
			var folder = fso.GetFolder(path);
		}
		catch (e) {
			return 0;
		}
		fc = new Enumerator(folder.files);

		for (; !fc.atEnd(); fc.moveNext()) {
			this.AllSize += fc.item().size;
		}
	},
	getFolders: function(path) {
		try {
			var folder = fso.GetFolder(path);
		}
		catch (e) {
			return 0;
		}
		this.fc = new Enumerator(folder.SubFolders);
		this.TempFolderList = new Array();

		this.getFolderOne();

	},
	getFolderOne: function() {
		if (!this.fc.atEnd()) {
			try {
				this.AllSize += fc.item().size;
			}
			catch (e) {
				this.TempFolderList[this.TempFolderList.length] = this.fc.item().path
			}
			this.fc.moveNext();
			window.setTimeout("loadInnerFolder.getFolderOne()", 10);
		}
		else {
			this.PandingFolder.shift();
			this.PandingFolder = this.PandingFolder.concat(this.TempFolderList);
			window.setTimeout("loadInnerFolder.getAll()", 10);
		}
	},
	showSize: function() {
		this.ide.childNodes.item(tableContent.itemSize).innerText = fig2Words(this.AllSize);
		this.ide.childNodes.item(tableContent.itemStatus).innerText = "loading...";
		table1.rows[table1.rows.length - 1].cells[tableContent.itemSize].innerText = fig2Words(this.AllSize + Total_Size);
	},
	complete: function() {
		this.ide.childNodes.item(tableContent.itemSize).innerText = fig2Words(this.AllSize);
		this.ide.childNodes.item(tableContent.itemStatus).innerText = "Ready";
		Total_Size += this.AllSize;
		table1.rows[table1.rows.length - 1].cells[tableContent.itemSize].innerText = fig2Words(Total_Size);
		this.FuncStart = false;
	}
}

var getFldStage = 0;
var getFldCount = 0;
var getFldCaller = null;
var folderLength = 0;
function getFolders(Caller) {
	if (Caller) getFldCaller = getFolders.caller;
	switch (getFldStage) {
		case 0:
			waitMsg.innerText = "Geting Folders...";
			Process_Win.reset();
			getFldCount = 0;
			getFldStage++;
			break;

		case 1:
			try {
				var folder = fso.GetFolder(loc_path.value);
				fc = new Enumerator(folder.SubFolders);
				folderLength = 0;
				for (; !fc.atEnd(); fc.moveNext()) folderLength++;
				fc = new Enumerator(folder.SubFolders);
			}
			catch (e) {
				ErrMsg(e);
			}
			getFldStage++;
			break;

		case 2:
			/* Get Folders List */
			StartTime(getNow());
			while (!fc.atEnd()) {
				try {
					var s = fc.item();
					files[files.length] = new fl_fld(s.name, s.attributes, s.size, s.type, true);
					if (size_max == null && size_min == null) { size_max = s.size; size_min = s.size; }
					else {
						if (size_max < s.size) size_max = s.size;
						if (size_min > s.size) size_min = s.size;
					}
					Total_Size += s.size;
				}
				catch (e) {
					files[files.length] = new fl_fld(s.name, s.attributes, e.description, "File Folder", true);
				}
				fc.moveNext();
				getFldCount++;
				waitMsg.innerText = "Geting Folders(" + getFldCount + ")...";
				Process_Win.Show_Progress(getFldCount * 100 / folderLength);
				if (getUsedTime(getNow()) > 500) break;
				//				if(getFldCount%1==0) break;
			}
			if (fc.atEnd()) getFldStage++;
			break;

		default: getFldStage = 0; getFldCount = 0; window.setTimeout("getFldCaller()", 500); return; break;
	}

	if (!Err) { window.setTimeout("getFolders()", 500); }
}

var getFileStage = 0;
var getFilesCount = 0;
var getFilesCaller = null;
var fileLength = 0;
function getFiles(Caller) {
	if (Caller) getFilesCaller = getFiles.caller;
	switch (getFileStage) {
		case 0:
			waitMsg.innerText = "Geting Files...";
			Process_Win.reset();
			getFilesCount = 0;
			getFileStage++;
			break;

		case 1:
			try {
				var folder = fso.GetFolder(loc_path.value);
				fc = new Enumerator(folder.files);
				for (; !fc.atEnd(); fc.moveNext()) fileLength++;
				fc = new Enumerator(folder.files);
			}
			catch (e) {
				ErrMsg(e);
			}
			getFileStage++;
			break;

		case 2:
			/* Get Files List */
			StartTime(getNow());
			while (!fc.atEnd()) {
				try {
					s = fc.item();
					files[files.length] = new fl_fld(s.name, s.attributes, s.size, s.type, false);
					if (size_max == null && size_min == null) { size_max = s.size; size_min = s.size; }
					else {
						if (size_max < s.size) size_max = s.size;
						if (size_min > s.size) size_min = s.size;
					}
					Total_Size += s.size;
				}
				catch (e) {
					files[files.length] = new fl_fld(s.name, s.attributes, e.description, e.description, false);
				}
				fc.moveNext();
				getFilesCount++;
				waitMsg.innerText = "Geting Files(" + getFilesCount + ")...";
				Process_Win.Show_Progress(getFilesCount * 100 / fileLength);
				if (getUsedTime(getNow()) > 500) break;
				//				if(getFilesCount%10==0) break;
			}
			if (fc.atEnd()) getFileStage++;
			break;

		default: getFileStage = 0; getFilesCount = 0; window.setTimeout("getFilesCaller()", 500); return; break;
	}

	if (!Err) { window.setTimeout("getFiles()", 500); }
}


var crtNewStage = 0;
var crtNewCount = 0;
var CrtNewCaller = null;
function createNew(Caller) {
	if (Caller) CrtNewCaller = createNew.caller;
	switch (crtNewStage) {
		case 0:
			waitMsg.innerText = "Creating New List...";
			Process_Win.reset();
			crtNewCount = 0;
			crtNewStage++;
			break;

		case 1:
			StartTime(getNow());
			while (crtNewCount < files.length) {
				with (table1.insertRow()) {
					if (!isNaN(files[crtNewCount].size)) style.backgroundColor = getColor(files[crtNewCount].size, size_max, size_min);
					className = "out";
					onselectstart = selectTable;
					onmouseover = mouseOver;
					onmouseout = mouseOut;
					ondblclick = openFld;
					//onclick = mouseClick;
					oncontextmenu = showContext;
					with (insertCell()) {
						innerHTML = "<input type=checkbox>";
					}
					with (insertCell()) {
						innerText = files[crtNewCount].name;
					}
					with (insertCell()) {
						style.fontFamily = "Consolas, Courier New";
						innerText = figAttr(files[crtNewCount].attribute);
					}
					with (insertCell()) {
						innerText = files[crtNewCount].type;
					}
					with (insertCell()) {
						style.textAlign = "right";
						if (!isNaN(files[crtNewCount].size)) {
							innerText = fig2Words(files[crtNewCount].size);
						}
						else {
							innerText = files[crtNewCount].size;
						}
					}
					with (insertCell()) {
						style.textAlign = "center";
						if (!isNaN(files[crtNewCount].size)) innerText = "Ready";
						else innerHTML = "<SPAN style=\"cursor:hand; color:blue;\" onclick=\"event.cancelBubble = true; loadInnerFolder.start();\" onmouseover=\"this.style.textDecoration=\'underline\'\" onmouseover=\"this.style.textDecoration=\'none\'\">Click To Load</SPAN>"
					}
				}
				crtNewCount++;
				table1.rows[crtNewCount].isFolder = files[crtNewCount - 1].isFolder;
				Process_Win.Show_Progress(crtNewCount * 100 / files.length);
				if (getUsedTime(getNow()) > 500) break;
				//				if(crtNewCount%10==0) break;
			}
			if (!(crtNewCount < files.length)) crtNewStage++;
			waitMsg.innerText = "Creating New List(" + (Math.floor(crtNewCount * 100 / files.length)) + "%)...";
			break;
		case 2:

			/* Create Total Row */
			with (table1.insertRow()) {
				with (insertCell()) {
				}
				with (insertCell()) {
					innerText = "Total Size";
					style.fontWeight = "bold";
				}
				with (insertCell()) {
				}
				with (insertCell()) {
				}
				with (insertCell()) {
					style.textAlign = "right";
					innerText = fig2Words(Total_Size);
					style.fontWeight = "bold";
				}
			}
			crtNewStage++;
			break;

		default: crtNewStage = 0; crtNewCount = 0; window.setTimeout("CrtNewCaller()", 500); return; break;
	}
	window.setTimeout("createNew()", 500);
}

var shrtListStage = 0;
var shortCaller = null;
function shortList(Caller) {
	if (Caller) shortCaller = shortList.caller;
	switch (shrtListStage) {
		case 0:
			waitMsg.innerText = "Shorting...";
			shrtListStage++;
			break;

		case 1:
			StartTime(getNow());
			var shortI = 0, shortFilesLength = files.length;
			for (var i = 0; i < files.length - 1; i++) {
				for (var j = i + 1; j < files.length; j++) {

					if (isNaN(files[i].size)) // Testing that size is number
					{
						var temp = files[j];
						files[j] = files[i];
						files[i] = temp;
					}

					if (files[j].size > files[i].size) {
						var temp = files[j];
						files[j] = files[i];
						files[i] = temp;
					}
					shortI++;
				}
			}
			//alert(getUsedTime(getNow()));
			shrtListStage++;
			break;

		default: shrtListStage = 0; window.setTimeout("shortCaller()", 100); return files; break;
	}
	shortList();
}

var ClrRecordStage = 0;
var ClrCaller = null;
function ClearRecord(Caller) {
	if (Caller) ClrCaller = ClearRecord.caller;
	switch (ClrRecordStage) {
		case 0:
			waitMsg.innerText = "Clearing Old Record...";
			break;

		case 1:
			while (table1.rows.length > 1) {
				table1.deleteRow();
			}
			break;

		default: ClrRecordStage = 0; window.setTimeout("ClrCaller()", 100); return; break;
	}
	ClrRecordStage++;
	window.setTimeout("ClearRecord()", 100); ;
}

var Real_path;
function mouseClick() {
	var ide = event.srcElement;
	while (ide.tagName != "TR") {
		ide = ide.parentElement;
	}
	if (ide.isFolder) loc_path.value = Real_path + ide.childNodes.item(0).innerText + "\\";
}

var selectedTR = null;
var selectedColor = null;

function mouseOut() {
	if (selectedTR != null) {
		selectedTR.style.backgroundColor = selectedColor;
	}
	selectedTR = null;
}

function mouseOver() {
	if (selectedTR != null) {
		selectedTR.style.backgroundColor = selectedColor;
	}
	selectedTR = null;
	var ide = event.srcElement;
	while (ide.tagName != "TR") {
		ide = ide.parentElement;
	}
	selectedColor = ide.style.backgroundColor;
	ide.style.backgroundColor = "#cccccc";
	selectedTR = ide;
}

function ErrMsg(e) {
	getFileStage = 0;
	getFldStage = 0;
	loadStage = 0;
	crtNewStage = 0;
	shrtListStage = 0;
	ClrRecordStage = 0;
	getFldCount = 0;
	getFilesCount = 0;

	Err = true;

	waitBox_showHide('hide');
	alert((e) + "\n" + (e.number & 0xFFFF) + "\n" + e.description);
}

function goBack() {
	var path = loc_path.value.slice(0, -1).split("\\");
	var pathString = "";
	var i = 0;
	do {
		pathString += path[i] + "\\";
		i++;
	} while (i < path.length - 1);
	loc_path.value = pathString;
	load();
}

function setWaitPosition() {
	waitBox.style.pixelTop = document.body.scrollTop + 200;
}

///////////////////////////

function ShowMemory() {
	var DrvName = Real_path.split("\\").slice(0, 1);

	toMemory = ShowUsedSpace(DrvName) * 100 / ShowTotalSpace(DrvName);
	CountShowMem = 0;
	_ShowMemory();
}

var CurrentMemory = 0;
var toMemory = 0;
var CountShowMem = 0;
function _ShowMemory() {
	CountShowMem++;
	memory_bar.style.width = (CurrentMemory + (toMemory - CurrentMemory) * CountShowMem / 50) + "%";

	if (CountShowMem < 50) window.setTimeout("_ShowMemory()", 10);
	else CurrentMemory = toMemory;
}

//////////////////////////
function selectTable() {
	event.returnValue = false;
}
//////////////////////////
// Context Menu //
var PathContext = null;
var ContextRow = null;
function showContext() {
	with (ContextFile) {
		with (style) {
			pixelLeft = event.x + document.body.scrollLeft + ((event.x + document.body.scrollLeft) > (document.body.offsetWidth - offsetWidth - 20) ? (-offsetWidth) : 0);
			pixelTop = event.y + document.body.scrollTop + ((event.y + document.body.scrollTop) > (document.body.offsetHeight - offsetHeight - 10) ? (-offsetHeight) : 0);
			visibility = "visible";
		}
	}

	overLayer.style.visibility = 'visible';
	overLayer.style.width = document.body.scrollWidth;
	overLayer.style.height = document.body.scrollHeight;

	var ide = event.srcElement;
	while (ide.tagName != "TR") {
		ide = ide.parentElement;
	}
	ContextRow = ide;
	PathContext = Real_path + ide.childNodes.item(1).innerText;

	if (ide.isFolder) {
		FileOpen.disabled = false;
		FileExplorer.disabled = false;
	}
	else {
		if (checkExt(PathContext)) FileOpen.disabled = false;
		else FileOpen.disabled = true;

		FileExplorer.disabled = true;
	}

	event.returnValue = false;
}

function openFld() {
	var ide = event.srcElement;
	while (ide.tagName != "TR") {
		ide = ide.parentElement;
	}
	ContextRow = ide;
	if (ide.isFolder) { FileOpen.disabled = false; FileExplorer.disabled = false; }
	else { FileOpen.disabled = true; FileExplorer.disabled = true; }

	PathContext = Real_path + ide.childNodes.item(1).innerText;
	OpenFolder();
}


function FolderExplorer() {
	overLayer.style.visibility = 'hidden';
	ContextFile.style.visibility = 'hidden';
	overLayer.style.width = 0;
	overLayer.style.height = 0;
	try {
		window.open(PathContext);
	}
	catch (e) {

	}
}

function OpenContainF() {
	overLayer.style.visibility = 'hidden';
	ContextFile.style.visibility = 'hidden';
	overLayer.style.width = 0;
	overLayer.style.height = 0;
	try {
		window.open(Real_path);
	}
	catch (e) {

	}
}

function OpenFolder() {
	overLayer.style.visibility = 'hidden';
	ContextFile.style.visibility = 'hidden';
	overLayer.style.width = 0;
	overLayer.style.height = 0;
	if (ContextRow.isFolder) {
		loc_path.value = PathContext + "\\";
		PathContext = null;
		load();
	}
	else {
		openFile();
	}
}

function deleteItem() {
	overLayer.style.visibility = 'hidden';
	ContextFile.style.visibility = 'hidden';
	overLayer.style.width = 0;
	overLayer.style.height = 0;
	try {
		if (FileOpen.disabled) {
			//File
			if (!confirm("Are you Sure to delete file \"" + PathContext + "\"?")) return;
			var f = fso.GetFile(PathContext);
			f.Delete();
		}
		else {
			//Folder
			if (!confirm("Are you Sure to delete folder \"" + PathContext + "\"?")) return;
			fso.DeleteFolder(PathContext);
		}

		for (var i = 0; i < table1.rows.length; i++) {
			if (ContextRow == table1.rows[i]) { table1.deleteRow(i); break; }
		}
	}
	catch (e) {
		alert("Error: " + e.message);
	}
	ShowMemory();
}

function init() {
	Process_Win = new progressBar("Process_Div", { width: 300 }, "Process_Win");
	init_size();
	init_drive();
}
////////////////////////////////
//Progress////////
////////////////////////////////
var Process_Win = null;
var progressBar = function(Progress_perent, options, func_id) {	// Progress Bar Create function

	this.options = Object.extend({ width: 300, height: 15, left: 10, top: 10 }, options || {});
	this.Progress_Perent_ID = Progress_perent;
	this.ID = func_id;
	if (!document.all[Progress_perent].all["Progress_Bar"]) {
		var Progress_Bar_String = "";
		Progress_Bar_String += "<DIV ID=Progress_Bar STYLE=\"MARGIN-TOP:" + this.options.top + "; MARGIN-LEFT:" + this.options.left + "; WIDTH:" + this.options.width + "; HEIGHT:" + this.options.height + "\">";
		Progress_Bar_String += "<DIV><TABLE CELLPADDING=0 CELLSPACING=0><TR><TD><IMG STYLE=\"WIDTH:2; HEIGHT:" + this.options.height + "\" SRC=\"Graphics/Progress_Back_Left.png\"></TD><TD><IMG STYLE=\"WIDTH:" + (this.options.width - 4) + "; HEIGHT:" + this.options.height + "\" SRC=\"Graphics/Progress_Back_Middle.png\"></TD><TD><IMG STYLE=\"WIDTH:2; HEIGHT:" + this.options.height + "\" SRC=\"Graphics/Progress_Back_Right.png\"></TD></TR></TABLE></DIV>";
		Progress_Bar_String += "<DIV STYLE=\"MARGIN-TOP:-" + (this.options.height - 1) + "; MARGIN-LEFT:1;\"><TABLE CELLPADDING=0 CELLSPACING=0><TR><TD><IMG ID=Progress_Fill STYLE=\"HEIGHT:" + (this.options.height - 2) + "; WIDTH:0;\" SRC=\"Graphics/Progress_Bar.png\"></TD></TR></TABLE></DIV>";
		Progress_Bar_String += "</DIV>";

		document.all[Progress_perent].insertAdjacentHTML("BeforeEnd", Progress_Bar_String);
	}

	this.progress = 0;
	this.change_progress_size = function(Precent) {

		if (Precent > 500) {
			//	alert("Error : Progress is greter than 100");
			Precent = 500
		}
		document.all[this.Progress_Perent_ID].all.Progress_Fill.style.pixelWidth = Precent * (this.options.width - 2) / 500;
		this.progress = Precent;
	}

	this.prog_size = 0;
	this.prog_size_start = 0;
	this.prog_size_to = 0;
	this.change_speed_t = 50;
	this.change_speed_d = 0;
	this.start = 0;
	this.high_light_bar = null;

	this.reset = function() {

		this.change_progress_size(0);
		this.prog_size = 0;
		this.prog_size_to = 0;
		this.change_speed_t = 0;
		this.change_speed_d = 0;
		this.start = 0;
	}

	this.Show_Progress = function(p) {

		this.prog_size_to = p * 5;
		this.prog_size_start = this.prog_size;

		if (!this.start) {
			this.start = 1;
			this.change();
		}
	}

	this.change = function() {

		if (this.prog_size < this.prog_size_to) {
			this.prog_size += this.change_speed_d;
			if (this.prog_size > this.prog_size_to) this.prog_size = this.prog_size_to;
			this.change_progress_size(this.prog_size);
			this.set_speed();
			window.setTimeout(this.ID + ".change()", 10);
		}
		else this.start = 0;
	}

	this.set_speed = function() {

		this.change_speed_d = (this.prog_size_to - this.prog_size_start) / 10;
	}


}

Object.extend = function extend(destination, source) {
	for (var property in source)
		destination[property] = source[property];
	return destination;
}

function checkAll(ide) {
	for (var i = 1; i < table1.rows.length - 1; i++) {
		table1.rows(i).cells(tableContent.checkBox).children(0).checked = ide.checked;
	}
}

function deleteMarked() {
	overLayer.style.visibility = 'hidden';
	ContextFile.style.visibility = 'hidden';
	overLayer.style.width = 0;
	overLayer.style.height = 0;
	if (!confirm("Are you Sure to delete Marked?")) return;
	var ErrorFound = 0;

	for (var i = 1; i < table1.rows.length - 1; i++) {
		if (table1.rows(i).cells(tableContent.checkBox).children(0).checked) {
			var delItemName = Real_path + table1.rows(i).cells(tableContent.itemName).innerText;
			if (table1.rows(i).isFolder) // check for folder item
			{
				try {
					fso.DeleteFolder(delItemName);
					table1.deleteRow(i);
					i--;
				}
				catch (e) {
					ErrorFound++;
				}
			}
			else {
				try {
					var f = fso.GetFile(delItemName);
					f.Delete();
					table1.deleteRow(i);
					i--;
				}
				catch (e) {
					ErrorFound++;
				}
			}
		}
	}

	if (ErrorFound) alert(ErrorFound + " Item is not Deleted.");
	ShowMemory();
}

///////////////////////////
///Drive List Object
///////////////////////////
function ShowDriveList() {
	DriveList.style.visibility = "visible";
	DriveList.focus();
	for (var i = 0; DriveList.rows.length > 0; i++)
		DriveList.deleteRow(0);
	var fso, s, n, e, x;
	fso = new ActiveXObject("Scripting.FileSystemObject");
	e = new Enumerator(fso.Drives);
	s = "";
	for (; !e.atEnd(); e.moveNext()) {
		x = e.item();
		if (x.DriveType == 3);
		else if (x.IsReady) {
			n = x.VolumeName;
			with (DriveList.insertRow(DriveList.rows.length)) {
				className = "out";
				onmouseover = mouseOver;
				onmouseout = mouseOut;
				onclick = SelectDrive;
				with (insertCell(0)) {
					innerText = x + "\\";
				}
			}
			DriveList.rows(DriveList.rows.length - 1).DrivePath = x + "\\";
		}
	}
}

function SelectDrive() {
	var RowElement = event.srcElement;
	while (RowElement.tagName != "TR")
		RowElement = RowElement.parentElement;
	loc_path.value = RowElement.DrivePath;
}

function checkPath() {
	try {
		fso.GetFolder(loc_path.value);
	}
	catch (e) {
		alert(e.description);
		return false;
	}
	return true;
}

//////////////////////////////////////
/////////////Open File////////////////
//////////////////////////////////////

var FileExtType = {
	"doc": "explorer",
	"docx": "explorer",
	"jpg": "explorer",
	"gif": "explorer",
	"bmp": "explorer",
	"png": "explorer",
	"ini": "notepad",
	"xls": "excel",
	"xlsx": "excel",
	"mp3": "wmplayer",
	"zip": "explorer"
}

function checkExt(path) {
	var fName = fso.GetFile(path).name;
	var Ext = "";
	for (var i = fName.length - 1; i >= 0; i--) {
		if (fName.substr(i, 1) == ".") break;
		Ext = fName.substr(i, 1) + Ext;
	}
	if (FileExtType[Ext]) return FileExtType[Ext];
	else return false;
}

function openFile() {
	//	if(checkExt(PathContext))
	//		Run(checkExt(PathContext)+' "'+PathContext+'"');

	Run("explorer" + ' "' + PathContext + '"');
}

function Run(strPath) {

	try {
		var objShell = new ActiveXObject("wscript.shell");
		objShell.Run(strPath);
		objShell = null;
	} //EO try  

	catch (e) {
		alert(e + "\nError Code : " + (e.number & 0xFFFF) + "\nError Massage : " + e.description, "OK");
	}

} //EO function

//////////////////////////////////////