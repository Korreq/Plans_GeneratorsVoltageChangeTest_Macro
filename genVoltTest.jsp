//Loading kdm model file
ReadDataKDM( "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\MODEL_ZW.kdm" );

var homeFolder = "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\LinearProgTest";
var tmpFile = homeFolder + "\\tmp.bin", tmpOgFile = homeFolder + "\\tmpOrg.bin";

// var nodeNameKeyWord = "DBN";
// var generatorNameKeyWord = "OPL_";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );


SaveTempBIN( tmpOgFile );

//Changeing iteration precision of UzG to 0.0001
Calc.EpsUg = 0.0001;



//var gen = getGenerators( 90, 1, 1 );
//var nd = getNodes( 90, 1, 1 );



var nodes = [], generators = [];

var oNodesFG = [];

var baseGensReacPow = [], baseNodesVolt = [], baseGenNodesPow = [];

var result = [];

for( var i = 1; i < Data.N_Nod; i++ ){

  var n = NodArray.Get( i );

  if( n.Area === 1 && n.St > 0 && n.Name.charAt(0) != 'Y' && n.Vn >= 90 ){ 
    
    nodes.push( n );
    
    baseNodesVolt.push( n.Vi );
  }
}

for( var i = 1; i < Data.N_Gen; i++  ){

  var g = GenArray.Get( i );

  var n = NodArray.Get( g.NrNod );

  if( g.Qmin !== g.Qmax && g.St > 0 && n.Area === 1 && n.Name.charAt(0) == 'Y' ){

    generators.push( [ g, n ] );

    baseGensReacPow.push( g.Qg );

  
    
    baseGenNodesPow.push( n.Vs );
  }
 
}

var file1 = fso.CreateTextFile( homeFolder + "\\ResultsG.csv" );
var file2 = fso.CreateTextFile( homeFolder + "\\ResultsN.csv" );

CalcLF();

//Write headers to files

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

SaveTempBIN( tmpFile )

for( i in generators ){
  
  var g = generators[ i ][ 0 ] ;
    
  var n = generators[ i ][ 1 ];
    
  if( g.TrfName.charAt( g.TrfName.length - 1 ) == 'A' ){

    var tName =  g.TrfName.slice( 0, -1 );

    var l = 'A';

    var t = null;

    while( true ){
      try{
        
        t = TrfArray.Find( tName + l );
        
        t.Typ = 11;
      }
      catch( e ){
        break;
      }
      //if( t === null) break;
      

      l = String.fromCharCode ( l.charCodeAt( 0 ) + 1 );
    }

  }
  else{
   
    try{
    TrfArray.Find( g.TrfName ).Typ = 11;
    }
    catch( e ){}
  }

  var value = 1
  
  n.Vs += value;

  CalcLF();
  
  file1.Write( g.Name + ";" + roundTo( baseGenNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );

  for( j in generators ){

    file1.Write( roundTo( generators[ j ][ 0 ].Qg, 2 ) + ";" );
  }
  
  file1.WriteLine("");

  file2.Write( g.Name + ";" + roundTo( baseGenNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );
  
  for( j in nodes ){

    file2.Write( roundTo( nodes[ j ].Vi, 2 ) + ";" );
  }

  file2.WriteLine("");

  ReadTempBIN( tmpFile );
}

ReadTempBIN( tmpOgFile );

fso.DeleteFile( tmpFile );
fso.DeleteFile( tmpOgFile );

file1.Close();
file2.Close();

function roundTo( value, precision ){

  return Math.round( value * ( 10 * precision ) ) / ( 10 * precision ) ;
}

function stringContainsWord( string, word ){
  
  var j = 0;

  for( var i = 0; i < string.length; i++ ){
  
    j = ( string.charAt( i ) === word.charAt( j ) ) ? j + 1 : 0;
  
    if( j === word.length ) return true;
  }
  
  return false;
}

function elementInArrayByName( array, elementName ){

   for( i in array ){
     
    e = array[ i ];
    
    if(e.Name === elementName) return true;
  }

  return false;
}

function getGeneratorNode( generator ){

  var n;

  if( generator.NodName.charAt(0) === "Y" ){

    var t = TrfArray.Find( generator.TrfName );

    n = NodArray.Find( t.EndName );

  }
  else{

    n = NodArray.Find( generator.NodName );

  }

  return n;

}


//Basic error thrower
function errorThrower( message, error ){
  
  MsgBox( message, 16, "Error" );
  throw error;
}

//Function returns filtred nodes array. Throws error if number of nodes in a model is less than 1.
//Takes ratedVoltage, areaId and status as filtration arguments 
function getNodes( ratedVoltage, areaId, status ){

  if( Data.N_Nod < 2 ) errorThrower( "No nodes in model", "No nodes in model, check model's nodes table" );

  var nodes = [];
  
  for( var i = 1; i < Data.N_Nod; i++ ){
  
    var node = NodArray.Get( i );
    
    if( node.Vn >= ratedVoltage && node.Area === areaId &&  node.St == status ) nodes.push( node ); 
  }

  return nodes;
}

//Function returns filtred generators array. Throws error if number of generators in a model is less than 1.
//Takes reactivePowerChange, areaId and status as filtration arguments 
function getGenerators( ratedVoltage, areaId, status ){

  if( Data.N_Gen < 2 ) errorThrower( "No generators in model", "No generators in model, check model's generators table" ); 

  var generators = [];

  for( var i = 1; i < Data.N_Gen; i++ ){
  
    var generator = GenArray.Get( i );
    //searching through built in NodArray to get generator's node  
    var genNode = NodArray.Get( generator.NrNod );
    
    if( genNode.Area === areaId && generator.St == status && genNode.Vn >= ratedVoltage ) generators.push( generator );
  }
  
  return generators;
}