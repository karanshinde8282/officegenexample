var fs = require('fs');
var path = require('path');


var officegen = require('./plugins');

var IMAGEDIR = __dirname + "/images/";
var OUTDIR = __dirname + "/output/";


//------------------------------------------------------------------------------------------------------------
var xlsx = officegen('xlsx');
sheet = xlsx.makeNewSheet ();
sheet.name = 'My Sheet';
sheet.setCell ( 'E7', 340 );
sheet.setCell ( 'G12', 'Hello World!' );
 
// Direct way: 
sheet.data[0] = [];
sheet.data[0][0] = 1;
sheet.data[0][1] = 2;
sheet.data[1] = [];
sheet.data[1][3] = 'abc';
sheet.setColumnWidth("G",25);
sheet.setColumnCenter("B");


sheet = xlsx.makeNewSheet ();
sheet.name = 'My new Sheet';

sheet.data[0] = [];
sheet.data[0][0] = 1;
sheet.data[1] = [];
sheet.data[1][3] = 'abc';
sheet.data[1][4] = 'More';
sheet.data[1][5] = 'Text';
sheet.data[1][6] = 'Here';
sheet.data[2] = [];
sheet.data[2][5] = 'abc';
sheet.data[2][6] = 900;
sheet.data[6] = [];
sheet.data[6][2] = 1972;
sheet.setCell('E8','=SUM(A1,A2)');
sheet.setCell ( 'E7', 340 );
sheet.setCell ( 'I1', -3 );
sheet.setCell ( 'I2', 31.12 );
sheet.setCell ( 'G102', 'Hello World!' );


var FILENAME = "test-xlsx1.xlsx";
var out = fs.createWriteStream(OUTDIR + FILENAME);
xlsx.generate(out);

//------------------------------------------------------------------------------------------------------------


