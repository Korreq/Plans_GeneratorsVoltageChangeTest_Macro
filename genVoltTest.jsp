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
for( var i = 1; i < Data.N_Gen; i++  ){

  var g = GenArray.Get( i );

  var n = NodArray.Get( g.NrNod );

  if( g.Qmin !== g.Qmax && g.St > 0 && n.Area === area && n.Name.charAt( nodeIndex ) == nodeChar ){

    generators.push( [ g, n ] );

    baseGensReacPow.push( g.Qg );

    baseGenNodesPow.push( n.Vs );
  }
 
}

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

  //Calculate power flow, if fails throw error 
  CPF();
  
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

//Basic error thrower
function errorThrower( message, error ){
  
  MsgBox( message, 16, "Error" );
  throw error;
}

//Calls built in power flow calculate function, throws error when it fails
function CPF(){

  if( CalcLF() != 1 ) errorThrower( "Power Flow calculation failed", -1 );
}

//Function takes conf object and depending on it's config creates folder in specified location. 
//Throws error if conf object is null and when folder can't be created
function createFolder( conf, fso ){

  var message = "Unable to load configuration";
  
  if( !conf ) errorThrower( message, message );
  
  var folder = conf.folderName;
  var folderPath = conf.homeFolder + "\\" + folder;
  
  if( !fso.FolderExists( folderPath ) ){
    
    try{ fso.CreateFolder( folderPath ); }
    
    catch( err ){ 
    
      errorThrower( "Unable to create folder", "Unable to create folder, check if you are able to create folders in that location" );
    }

  }
  
  folder += "\\";

  return folder;
}

//Function takes conf object and depending on it's config creates file in specified location.
//Also can create folder where results are located depending on configuration file 
//Throws error if conf object is null and when file can't be created
function createFile( fileNameEnd, conf, fso ){
  
  var message = "Unable to load configuration";
  if( !conf ) errorThrower( message, message );

  var file = null;
  
  var folder = ( conf.createResultsFolder == 1 ) ? createFolder( conf, fso ) : "";
  var timeStamp = ( conf.addTimestampToFile == 1 ) ? getCurrentDate() + "--" : "";
  var fileLocation = conf.homeFolder + "\\" + folder + timeStamp + conf.fileName + fileNameEnd + ".csv";
  
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( err ){ 
    
    errorThrower( "File arleady exists or unable to create it", "File arleady exists or unable to create it, check if you are able to create files in that location" );
  }

  return file;
} 

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