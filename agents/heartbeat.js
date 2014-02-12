////////////////////////////////////////////////////////////////////////////////////////
//
//	heartbeat.js - used to monitor and audit WINDOWS workstations to JSON files
//	and POST those JSON files to a heartbeat server
//	run on target machine or through GPO using cscript
//
//		cscript heartbeat.js http://server:port [logon/logoff/turnon/turnoff]
//
//	tested on Windows 7 workstations.
//
//	creates the following object

//	{
//		"Name":"STRING",
//		"Model":"STRING",
//		"Manufacturer":"STRING",
//		"ServiceTag":"STRING",
//		"SystemType":"STRING",
//		"TotalPhysicalMemory":NUM,
//		"OS":{
//			"Name":"STRING",
//			"Serial":"STRING"
//		},
//		"Software":[
//			{
//				"Name":"STRING",
//				"Vendor":"STRING",
//				"Version":"STRING",
//				"InstallDate":"STRING"
//			},
//		],
//		"UserName":"STRING"
//	}
//
//
//
////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
// Set variables

var heartServer = WScript.arguments(0);	
var mode = WScript.arguments(1);	
var shell = new ActiveXObject("WScript.Shell");
var fs = new ActiveXObject("Scripting.FileSystemObject");
var http = new ActiveXObject("Microsoft.XMLHTTP"); 
var AppData = shell.ExpandEnvironmentStrings("%AppData%")
var fileName = AppData + "\\heartbeat.json";
var wmi = GetObject("winmgmts://./root/cimv2");
var compObj = {};
var softwareObj = {};

////////////////////////////////////////////////////////////////////////////////////////
// set functions

function fixUser(string) //removes domain and slash from windows usernames
{
	string = string.split("\\");
    return string[string.length-1].toUpperCase();
};

function fixOSName(data, callback) //tidies up OS Naming styles and checks if OEM OS
{
	var oemReg = new RegExp("OEM");	
	data.Name = data.Name.replace(/MICROSOFT|SERVER|\(R\)|�/g,""); //remove unwanted
	data.Name = data.Name.replace(/\s\s/g," "); //remove double space
	data.Name = data.Name.replace(/^\s+|\s+$/g,""); //trim	
	//test for OEM
	if(oemReg.test(data.Serial)){
		data.Name = data.Name+" OEM"
	};	
	callback(data);
}

function getOSInfo(callback) //gets OS information from WMI
{
	var osInfo = wmi.ExecQuery("SELECT SerialNumber,Caption FROM Win32_OperatingSystem");
	var enumItems = new Enumerator(osInfo);	
	for (;!enumItems.atEnd();enumItems.moveNext()) {
		var objItem = enumItems.item();
		var osObj = {};
		osObj.Serial = objItem.SerialNumber.toUpperCase();		
		osObj.Name = objItem.Caption.toUpperCase();
		fixOSName(osObj,function(data){
			compObj.OS = data;
		});
	}
	callback();
}

function getComputerInfo(callback) //gets Computer information from WMI
{
	var compInfo =  wmi.ExecQuery("SELECT Name,Model,Manufacturer,SystemType,TotalPhysicalMemory,UserName FROM Win32_ComputerSystem");
	var enumItems = new Enumerator(compInfo);	
	for (;!enumItems.atEnd();enumItems.moveNext()) {
		var objItem = enumItems.item();
		compObj.Name = objItem.Name.toUpperCase(); //I want uppercase
		compObj.Model = objItem.Model.toUpperCase();
		compObj.Manufacturer = objItem.Manufacturer.toUpperCase();
		compObj.SystemType = objItem.SystemType.toUpperCase();
		compObj.TotalPhysicalMemory = objItem.TotalPhysicalMemory;
		if(objItem.UserName == null){compObj.UserName = "Unknown"}else{
			
			compObj.UserName = fixUser(objItem.UserName); //remove domain
		}
	}
	callback();
}

function getServiceTag(callback) //gets machines serial/servicetag from WMI
{
	var svTag = wmi.ExecQuery("SELECT SerialNumber FROM Win32_BIOS");
	var enumItems = new Enumerator(svTag);	
	for (;!enumItems.atEnd();enumItems.moveNext()) {
		var objItem = enumItems.item();
		compObj.ServiceTag = objItem.SerialNumber.toUpperCase();
	}
	callback();
}

function getInstalledSoftware(callback) //use Windows Installer to search for products
{
	WScript.Echo("1");
	var installer = new ActiveXObject("WindowsInstaller.Installer");
	var products = installer.Products;
	var softwareArray = [];
	WScript.Echo("2");
	WScript.Echo(installer.ProductInfo(products.Item(1), "Publisher"));
	for(var i=0;i<products.Count;i++){
		var softObj = {};
		try {
			WScript.Echo(i +" / "+ products.Count);
			softObj.Name = "\""+installer.ProductInfo(products.Item(i),"InstalledProductName").toUpperCase()+"\"";
			softObj.InstallDate = "\""+installer.ProductInfo(products.Item(i), "InstallDate").toUpperCase()+"\"";
			softObj.Vendor = "\""+installer.ProductInfo(products.Item(i), "Publisher").toUpperCase()+"\"";
			softObj.Version = "\""+installer.ProductInfo(products.Item(i), "VersionString").toUpperCase()+"\"";
		}
		catch(e){	
		}			
		softwareArray.push(softObj);
	}
	callback(softwareArray);
}

function writeFile(softwareArray, mode, callback){
	var file = fs.CreateTextFile(fileName);
	
	file.WriteLine("{");
	file.WriteLine('"Name":"'+compObj.Name+'",');
	file.WriteLine('"Model":"'+compObj.Model+'",');
	file.WriteLine('"Manufacturer":"'+compObj.Manufacturer+'",');
	file.WriteLine('"ServiceTag":"'+compObj.ServiceTag+'",');
	file.WriteLine('"SystemType":"'+compObj.SystemType+'",');
	file.WriteLine('"TotalPhysicalMemory":'+compObj.TotalPhysicalMemory+',');
		file.WriteLine('"OS":{');			
			file.WriteLine('"Name":"'+compObj.OS.Name+'",');
			file.WriteLine('"Serial":"'+compObj.OS.Serial+'"');
		file.WriteLine("},");
	file.WriteLine('"Software":[');
	
	for(var i=0;i<softwareArray.length;i++){
		if(typeof softwareArray[i].Name === "undefined"){
		}else{
			file.Write("{");
			file.Write("\"Name\":"+softwareArray[i].Name+",");
			file.Write("\"Vendor\":"+softwareArray[i].Vendor+",");
			file.Write("\"Version\":"+softwareArray[i].Version+",");
			file.Write("\"InstallDate\":"+softwareArray[i].InstallDate+"");	
			file.WriteLine("}" + (i==softwareArray.length-1 ? '],': ','));
		}
	};
	
	file.WriteLine('"UserName":"'+compObj.UserName+'",');
	file.WriteLine('"Mode":"'+mode+'"');
	file.WriteLine("}");
	file.Close();
	callback();
}


function sendData(data){
	WScript.Echo("Sending Data");
	http.open("POST", heartServer, false);
	http.send(data);
};

function readData(){
	WScript.Echo("Reading Data");
	var file = fs.GetFile(fileName);
    var stream = file.OpenAsTextStream(1, -2);
	var data = stream.ReadAll();
	stream.Close();
	sendData(data)
};

function runScript(mode){
	getServiceTag(function(){
		getOSInfo(function(){
			getComputerInfo(function(){
				getInstalledSoftware(function(softwareArray){
					writeFile(softwareArray, mode, function(){
						readData();
					});
				});
			});
		});
	});
}

////////////////////////////////////////////////////////////////////////////////////////
// run 
// chooses mode based on mode varible (logon logoff etc...)

try{
	switch(mode){
		case "logon":
			shell.Run(runScript("logon"));
			break;
		case "logoff":
			shell.Run(runScript("logoff"));
			break;
		case "turnon":			
			shell.Run(runScript("turnon"));
			break;
		case "turnoff":			
			shell.Run(runScript("turnoff"));
			break;
		default:
			shell.Run(runScript("default"));
			break;
	}
}
catch(e){
};
