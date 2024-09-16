//TODO list
/*  
  
  configuraton file with file reading support  
  
  convert input file array to plans objects
  
  search function, that searches through generators, transformers and nodes?

  change output files to log change of transformers and generators and nodes

*/

//Location of folder where config file is located
var homeFolder = "C:\\Users\\lukas\\Documents\\Github\\Plans_GeneratorsVoltageChangeTest_Macro";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

//Initializing configuration object
var conf = iniConfigConstructor( homeFolder, fso );
var tmpFile = conf.homeFolder + "\\tmp.bin", tmpOgFile = conf.homeFolder + "\\tmpOrg.bin";

//Loading kdm model file and trying to save it as temporary binary file
ReadDataKDM( conf.modelPath + "\\" + conf.modelName + ".kdm" );
if( SaveTempBIN( tmpOgFile ) < 1 ) errorThrower( "Unable to create temporary file", "Unable to create temporary file, check if you are able to create files in homeFolder location" );

var time = getTime();

var nodes = [], generators = [];

var baseGensReacPow = [], baseNodesVolt = [], baseGenNodesPow = [];

//Getting variables values from config file 
var area = conf.area, voltage = conf.voltage, nodeIndex = conf.nodeIndex, nodeChar = conf.nodeChar;

//Setting power flow calculation settings with settings from config file
setPowerFlowSettings( conf );

//Calculate power flow, if fails throw error 
CPF();


//
//TODO: convert data from input file to array with objects
//

//Try to read file from location specified in configuration file, then make array from file and close the file
var inputFile = readFile( config, fso );
var inputArray = getInputArray( inputFile );
inputFile.close();
/*
//Fill node array with valid nodes and baseNodesVolt array with voltage of that nodes
for( var i = 1; i < Data.N_Nod; i++ ){

  var n = NodArray.Get( i );

  if( n.Area === area && n.St > 0 && n.Name.charAt( nodeIndex ) != nodeChar && n.Vn >= voltage ){ 
    
    nodes.push( n );
    
    baseNodesVolt.push( n.Vi );
  }
}

//Fill generators array with valid generators and it's connected node. 
//Also fill baseGensReacPow with that generators reactive power and baseGenNodesPow with connected nodes power
for( var i = 1; i < Data.N_Gen; i++ ){

  var g = GenArray.Get( i );

  var n = NodArray.Get( g.NrNod );

  if( g.Qmin !== g.Qmax && g.St > 0 && n.Area === area && n.Name.charAt( nodeIndex ) == nodeChar ){

    generators.push( [ g, n ] );

    baseGensReacPow.push( g.Qg );

    baseGenNodesPow.push( n.Vs );
  }
 
}
*/
//Create result files and folder with settings from a config file
var file1 = createFile( "G", conf, fso );
var file2 = createFile( "N", conf, fso );

//Write headers and base values for each generator/node to coresponding file 
file1.Write( "Generator;Old U_G;New U_G;" );
file2.Write( "Generator;Old U_G;New U_G;" );

var temp = "Base;X;X;";

for( i in generators ){

  file1.Write( generators[ i ][ 0 ].Name + ";" );

  temp += roundTo( baseGensReacPow[ i ], 2 ) + ";";
}

file1.WriteLine( "\n" + temp );

temp = "Base;X;X;";

for( i in nodes ){

  file2.Write( nodes[ i ].Name + ";" );

  temp += roundTo( baseNodesVolt[ i ], 2 ) + ";";
}

file2.WriteLine( "\n" + temp );

//Trying to save file before changes on transformators and  connected nodes
if( SaveTempBIN( tmpFile ) < 1 ) errorThrower( "Unable to create temporary file", "Unable to create temporary file, check if you are able to create files in homeFolder location" );

for( i in generators ){
  
  var g = generators[ i ][ 0 ], n = generators[ i ][ 1 ];
  
  //Check if generator has block transformer
  if( g.TrfName != "" ){

    //Find transformer and change it's type to 11 ( without regulation )
    var t = TrfArray.Find( g.TrfName );
    t.Typ = 11;
    
    //Check if transformer name's ends with A, indicates that there are more than 1 block transformers connected to generator
    if( g.TrfName.charAt( g.TrfName.length - 1 ) == 'A' ){
      
      var l = 'B';
      //Transformer name without last char
      var tName = g.TrfName.slice(0, -1);
      
      //As long as there are transformers with same name ending with next letter, then change it's type to 11 ( without regulation )
      while( true ){
        
        //Try to assign a transformer to variable and transformer type to 11 ( without regulation )
        try{
          
          t = TrfArray.Find( tName + l ); 
          t.Typ = 11;
        }
        
        //If t is null then exit while loop
        catch( e ){ break; }
        
        //changes l to next letter by adding 1 to it's char code
        l = String.fromCharCode ( l.charCodeAt( 0 ) + 1 );
      }

    }

  }

  //get set value from config file and add it to node's voltage
  var value = conf.value;
  n.Vs += value;

  //Calculate power flow, if fails try to load original model and throw error 
  if( CalcLF() != 1 ) saveErrorThrower( "Power Flow calculation failed", -1, tmpOgFile );

  //Write generator's name, it's base connected node power and new connected node power
  file1.Write( g.Name + ";" + roundTo( baseGenNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );

  //Write for each generator it's new reactive power
  for( j in generators ){

    file1.Write( roundTo( generators[ j ][ 0 ].Qg, 2 ) + ";" );
  }
  
  //Add end line character to file
  file1.WriteLine("");

  //Write generator's name, it's base connected node power and new connected node power
  file2.Write( g.Name + ";" + roundTo( baseGenNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );
  
  //Write for each node it's new voltage
  for( j in nodes ){

    file2.Write( roundTo( nodes[ j ].Vi, 2 ) + ";" );
  }

  //Add end line character to file
  file2.WriteLine("");

  //Load model without any changes to transformators
  ReadTempBIN( tmpFile );
}

//Loading original model
ReadTempBIN( tmpOgFile );

//Removing temporary binary files
fso.DeleteFile( tmpFile );
fso.DeleteFile( tmpOgFile );

//Closing result files
file1.Close();
file2.Close();

time = getTime() - time;
cprintf( "Time Elapsed: "+ time ); 

//Function uses JS Math.round, takes value and returns rounded value to specified decimals 
function roundTo( value, precision ){

  return Math.round( value * ( 10 * precision ) ) / ( 10 * precision ) ;
}

//Set power flow settings using config file
function setPowerFlowSettings( config ){

  Calc.Itmax = config.maxIterations;
  Calc.EPS10 = config.startingPrecision;
  Calc.Eps = config.precision;
  Calc.EpsUg = config.uzGIterationPrecision;
  Calc.Met = config.method;
}

//Function adds loading bin file before throwing an error
function saveErrorThrower( message, error, binPath ){

  try{ ReadTempBIN( binPath ); }
  
  catch( e ){ MsgBox( "Couldn't load original model", 16 ) }

  errorThrower( message, error );
}

//Basic error thrower with error message window
function errorThrower( message ){
  
  MsgBox( message, 16, "Error" );
  throw message;
}

//Calls built in power flow calculate function, throws error if it fails
function CPF(){

  if( CalcLF() != 1 ) errorThrower( "Power Flow calculation failed" );
}

//Function takes config object and depending on it's config creates folder in specified location. 
//Throws error if config object is null and when folder can't be created
function createFolder( config, fso ){
  
  if( !config ) errorThrower( "Unable to load configuration" );
  
  var folder = config.folderName;
  var folderPath = config.homeFolder + folder;
  
  if( !fso.FolderExists( folderPath ) ){
    
    try{ fso.CreateFolder( folderPath ); }
    
    catch( err ){ errorThrower( "Unable to create folder" ); }
  }
  
  folder += "\\";

  return folder;
}

//Function takes config object and depending on it's config creates file in specified location.
//Also can create folder where results are located depending on configuration file 
//Throws error if config object is null and when file can't be created
function createFile( config, fso ){
 
  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;
  
  var folder = ( config.createResultsFolder == 1 ) ? createFolder( config, fso ) : "";
  var timeStamp = ( config.addTimestampToFile == 1 ) ? getCurrentDate() + "--" : "";
  var fileLocation = config.homeFolder + folder + timeStamp + config.resultFileName + ".txt";
  
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( err ){ errorThrower( "File already exists or unable to create it" ); }

  return file;
} 

//Function takes config object and depending on it reads file from specified location.
//Throws error if config object is null or when file can't be read 
function readFile( config, fso ){

  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;

  var fileLocation = config.inputFileLocation + config.inputFileName + "." + config.inputFileFormat;
  
  try{ file = fso.OpenTextFile( fileLocation, 1, false, 0 ); }

  catch( err ){ errorThrower( "Unable to find or open file" ); }

  return file;
}

//
//TODO update old config file with the one from arst 
//

/*

//Function uses built in .ini function to get it's settings from config file.
//Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ){
  
  var configFile = iniPath + "\\config.ini";

  if( !fso.FileExists( configFile ) ) errorThrower( "config.ini file not found" );

  //Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( configFile );

  var hFolder = ini.GetString( "main", "homeFolder", Main.WorkDir );
  
  //Declaring conf object and trying to fill it with config.ini configuration
  var conf = {
  
    //Main
    homeFolder: hFolder,
    modelName: ini.GetString( "main", "modelName", "model" ),
    modelPath: ini.GetString( "main", "modelPath", hFolder ),  
    safeMode: ini.GetBool( "main", "safeMode", 1 ),

    //Folder
    createResultsFolder: ini.GetBool( "folder", "createResultsFolder", 0 ),
    folderName: ini.GetString( "folder", "folderName", "folder" ),
    
    //Files
    addTimestampToFile: ini.GetBool( "files", "addTimestampToFile", 1 ),
    inputFileLocation: ini.GetString( "files", "inputFileLocation", hFolder ),
    inputFileName: ini.GetString( "files", "inputFileName", "input" ),
    inputFileFormat: ini.GetString( "files", "inputFileFormat", "txt" ),
    resultFileName: ini.GetString( "files", "rsultFileName", "log" ),
    roundingPrecision: ini.GetInt( "files", "roundingPrecision", 2 ),
    
    //Power Flow
    maxIterations: ini.GetInt( "power flow", "maxIterations", 300 ),
    startingPrecision: ini.GetDouble( "power flow", "startingPrecision", 10.00 ),
    precision: ini.GetDouble( "power flow", "precision", 1.00 ),
    method: ini.GetInt( "power flow", "method", 1 )
  };
  
  //Overwriting config.ini file
  //Main
  ini.WriteString( "main", "homeFolder", conf.homeFolder );
  ini.WriteString( "main", "modelName", conf.modelName );
  ini.WriteString( "main", "modelPath", conf.modelPath );
  ini.WriteBool( "main", "safeMode", conf.safeMode );
  
  //Folder
  ini.WriteBool( "folder", "createResultsFolder", conf.createResultsFolder );
  ini.WriteString( "folder", "folderName", conf.folderName );
    
  //Files
  ini.WriteBool( "files", "addTimestampToFile", conf.addTimestampToFile );
  ini.WriteString( "files", "inputFileLocation", conf.inputFileLocation );
  ini.WriteString( "files", "inputFileName", conf.inputFileName );
  ini.WriteString( "files", "inputFileFormat", conf.inputFileFormat );
  ini.WriteString( "files", "resultFileName", conf.resultFileName );
  ini.WriteInt( "file", "roundingPrecision", conf.roundingPrecision );
    
  //Power Flow
  ini.WriteInt( "power flow", "maxIterations", conf.maxIterations );
  ini.WriteDouble( "power flow", "startingPrecision", conf.startingPrecision );
  ini.WriteDouble( "power flow", "precision", conf.precision );
  ini.WriteInt( "power flow", "method", conf.method );
 
  return conf;
}



*/

//Function uses built in .ini function to get it's settings from config file.
//Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ){
  
  var confFile = iniPath + "\\config.ini";

  if( !fso.FileExists( confFile ) ) errorThrower( "config.ini file not found", "Config file error, make sure your file location has config.ini file" );

  //Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( confFile );

  var hFolder = ini.GetString( "main", "homeFolder", Main.WorkDir );
  
  //Declaring conf object and trying to fill it with config.ini configuration
  var conf = {
  
    //Main
    homeFolder: hFolder,
    modelName: ini.GetString( "main", "modelName", "model" ),
    modelPath: ini.GetString( "main", "modelPath", hFolder ),  
    
    //Variable
    area: ini.GetInt( "variable", "area", 1 ),
    voltage: ini.GetInt( "variable", "voltage", 0 ),
    nodeIndex: ini.GetInt( "variable", "nodeIndex", 0 ),
    nodeChar: ini.GetString( "variable", "nodeChar", 'Y' ),
    value: ini.GetInt( "variable", "value", 1 ),

    //Folder
    createResultsFolder: ini.GetBool( "folder", "createResultsFolder", 0 ),
    folderName: ini.GetString( "folder", "folderName", "folder" ),
    
    //File
    addTimestampToFile: ini.GetBool( "file", "addTimestampToFile", 1 ),
    fileName: ini.GetString( "file", "fileName", "result" ),
    roundingPrecision: ini.GetInt( "file", "roundingPrecision", 2 ),
    
    //Power Flow
    maxIterations: ini.GetInt( "power flow", "maxIterations", 300 ),
    startingPrecision: ini.GetDouble( "power flow", "startingPrecision", 10.00 ),
    precision: ini.GetDouble( "power flow", "precision", 1.00 ),
    uzGIterationPrecision: ini.GetDouble( "power flow", "uzGIterationPrecision", 0.001 ),
    method: ini.GetInt( "power flow", "method", 1 )
  };
  
  //Overwriting config.ini file
  //Main
  ini.WriteString( "main", "homeFolder", conf.homeFolder );
  ini.WriteString( "main", "modelName", conf.modelName );
  ini.WriteString( "main", "modelPath", conf.modelPath );

  //Variable
  ini.WriteInt( "variable", "area", conf.area );
  ini.WriteInt( "variable", "voltage", conf.voltage );
  ini.WriteInt( "variable", "nodeIndex", conf.nodeIndex );
  ini.WriteString( "variable", "nodeChar", conf.nodeChar );
  ini.WriteInt( "variable", "value", conf.value );

  //Folder
  ini.WriteBool( "folder", "createResultsFolder", conf.createResultsFolder );
  ini.WriteString( "folder", "folderName", conf.folderName );
    
  //File
  ini.WriteBool( "file", "addTimestampToFile", conf.addTimestampToFile );
  ini.WriteString( "file", "fileName", conf.fileName );
  ini.WriteInt( "file", "roundingPrecision", conf.roundingPrecision );
    
  //Power Flow
  ini.WriteInt( "power flow", "maxIterations", conf.maxIterations );
  ini.WriteDouble( "power flow", "startingPrecision", conf.startingPrecision );
  ini.WriteDouble( "power flow", "precision", conf.precision );
  ini.WriteDouble( "power flow", "uzGIterationPrecision", conf.uzGIterationPrecision );
  ini.WriteInt( "power flow", "method", conf.method );
 
  return conf;
}

//Function gets file and takes each line into a array and after finding whiteline pushes array into other array
function getInputArray( file ){

  var array = [];

  while(!file.AtEndOfStream){

    var tmp = [], line = file.ReadLine();
      
    while( line != "" ){
     
      tmp.push( line.replace(/(^\s+|\s+$)/g, '') );
    
      if( !file.AtEndOfStream ) line = file.ReadLine();
      
      else break; 
    }
    
    array.push( tmp );
  }

  return array;
}

//Function takes current date and returns it in file safe format  
function getCurrentDate(){
  
  var current = new Date();
  
  var formatedDateArray = [ ( '0' + ( current.getMonth() + 1 ) ).slice( -2 ), ( '0' + current.getDate() ).slice( -2 ), 
  ( '0' + current.getHours() ).slice( -2 ), ( '0' + current.getMinutes() ).slice( -2 ), ( '0' + current.getSeconds() ).slice( -2 ) ];
  
  return current.getFullYear() + "-" + formatedDateArray[ 0 ] + "-" + formatedDateArray[ 1 ] + "--" + formatedDateArray[ 2 ] + "-" + formatedDateArray[ 3 ] + "-" + formatedDateArray[ 4 ];
}

function getTime(){
  
  var current = new Date();
  
  return current.getHours() * 3600 + current.getMinutes() * 60 + current.getSeconds();
}