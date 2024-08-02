//Loading kdm model file
ReadDataKDM( "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\MODEL.kdm" );

var homeFolder = "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\LinearProgTest";
var tmpFile = homeFolder + "\\tmp.bin", tmpOgFile = homeFolder + "\\tmpOrg.bin";

var nodeNameKeyWord = "DBN";
var generatorNameKeyWord = "OPL_";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

//Changeing iteration precision of UzG to 0.0001
Calc.EpsUg = 0.0001;

SaveTempBIN( tmpOgFile );

var nodes = [], generators = [];

var oNodesFG = [];

var baseGensReacPow = [], baseNodesVolt = [], baseGenNodesPow = [];

var result = [];

for( var i = 1; i < Data.N_Nod; i++ ){

  var n = NodArray.Get( i );

  if( stringContainsWord( n.Name, nodeNameKeyWord ) && n.St > 0 ) nodes.push( n );
}

for( var i = 1; i < Data.N_Gen; i++  ){

  var g = GenArray.Get( i );
  
  if( stringContainsWord( g.Name, generatorNameKeyWord ) && g.St > 0 ) generators.push( g );
}

for( i in generators ){
  
  var g = generators[ i ];
    
  var t = TrfArray.Find( g.TrfName );
  
  t.Typ = 11;
  
  CalcLF();
  
  baseGensReacPow.push( g.Qg );
  
  var nG = ( t.BegName === g.NodName )? t.EndName : t.BegName;
  
  for( j in nodes){
  
    if( nG === nodes[ j ].Name ){ 
      
      if( ! elementInArrayByName( oNodesFG, nG ) ){
      
        oNodesFG.push( nodes[ j ] ); 
        baseNodesVolt.push( nodes[ j ].Vi );
      }
    
    }
   
  }

}

var file = fso.CreateTextFile( homeFolder + "\\Results.csv" );

var value = 1

SaveTempBIN( tmpFile )


file.WriteLine( ";Start State;" );

var header = ";U_G;";
 
for( i in oNodesFG ){ header += oNodesFG[ i ].Name + ";" ; }

for( i in generators ){ header += generators[ i ].Name + ";" ; }

file.WriteLine( header );

for( i in generators ){

  var g = generators[ i ], n = NodArray.Find( g.NodName );
   
  baseGenNodesPow.push( n.Vs );
  file.Write( n.Name + ";" + roundTo( n.Vs, 2 ) + ";" );
  
  for( j in baseNodesVolt ){ file.Write( roundTo( baseNodesVolt[ j ], 2 ) + ";" ); }
  
  for( j in baseGensReacPow ){ file.Write( roundTo( baseGensReacPow[ j ], 2 ) + ";" ); }
  
  file.WriteLine( "" );
}

file.WriteLine( "\n;End State;\n" + header );

for( i in generators ){
   
  var g = generators[ i ];

  var n = NodArray.Find( g.NodName );
    
  n.Vs += value;
  
  CalcLF();
    
  file.Write( n.Name + ";" + roundTo( n.Vs, 2 ) + ";" );
  
  result.push( [] );
  result[i].push( n.Name ,roundTo( n.Vs - baseGenNodesPow[ i ], 2 )  );
  
  for( j in oNodesFG ){
  
    var nF = oNodesFG[ j ];
    
    file.Write( roundTo( nF.Vi, 2 ) + ";" );
    result[i].push( roundTo( nF.Vi - baseNodesVolt[ j ], 2 ) );
  }
  
  for( j in generators ){
    
    var g = generators[ j ];
  
    file.Write( roundTo( g.Qg, 2 ) + ";" );
    result[i].push( roundTo( g.Qg - baseGensReacPow[ j ] , 2 ) );
  
  }
  
  file.WriteLine( "" );
   
  ReadTempBIN( tmpFile );
}

file.WriteLine( "\n;Difference between Start state and End State;\n" + header );

for( i in result ){

  for( j in result[ i ] ){ file.Write( result[i][j] + ";" ); }

  file.WriteLine("");
}

ReadTempBIN( tmpOgFile );

fso.DeleteFile( tmpFile );
fso.DeleteFile( tmpOgFile );

file.Close();

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
