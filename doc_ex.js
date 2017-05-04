var fs = require('fs');
var path = require('path');


var officegen = require('./plugins');

var IMAGEDIR = __dirname + "/images/";
var OUTDIR = __dirname + "/output/";




//------------------------------------------------------------------------------------------------------------
var docx = officegen('docx');

var header = docx.getHeader().createP();
header.addText('asd');

var pObj = docx.createP();


pObj.options.align = 'left'; // Also 'right' or 'jestify'. 
pObj.addText('Simple text for showing addText() function'); pObj.addLineBreak();
pObj.addText(' with color', { color: '000088' }); pObj.addLineBreak();
pObj.addText(' and back color.', { color: '00ffff', back: '000088' });
pObj.addLineBreak(); pObj.addLineBreak(); pObj.addLineBreak();

var pObj = docx.createP();

pObj.addText('Bold + underline', { bold: true, underline: true });

var pObj = docx.createP({ align: 'center' });

pObj.addText('Center this text.');

var pObj = docx.createP();
pObj.options.align = 'right';

pObj.addText('Align this text to the right.');

var pObj = docx.createP({ backline: 'yellow' });

pObj.addText('Those two lines are in the same paragraph,');
pObj.addLineBreak();
pObj.addText('but they are separated by a line break.');


var pObj = docx.createP({align:'center'});
pObj.addLineBreak(); pObj.addLineBreak();
pObj.addImage(path.resolve(IMAGEDIR, 'image (1).jpg'), {cx: 300, cy: 200 });


var pObj = docx.createP ();

pObj.addText ( 'Since ',{font_size : 20 ,font_face:'Ravie'} );
pObj.addText ( 'asd', { back: '00ffff', shdType: 'pct12', shdColor: 'ff0000',font_size : 20 ,font_face:'Jokerman' } ); // Use pattern in the background.
pObj.addText ( ' you can do ' ,{font_size : 20 ,font_face:'Script MT'});
pObj.addText ( 'more cool ', { highlight: true ,font_size : 20 ,font_face:'Small Fonts'} ); // Highlight!
pObj.addText ( 'stuff!', { highlight: 'darkGreen' ,font_size : 20 ,font_face:'Lucida Calligraphy'} ); // Different highlight color.


docx.putPageBreak();

var pObj = docx.createP({align:'center'});


pObj.addLineBreak(); pObj.addLineBreak();
pObj.addImage(path.resolve(IMAGEDIR, 'image (2).jpg'), { cx: 500, cy: 500 });

pObj.addLineBreak();
var pObj = docx.createP();
pObj.addText('Fonts face only.', { font_face: 'Arial' });
pObj.addText(' Fonts face and size.', { font_face: 'JOKER', font_size: 25 });
pObj.addLineBreak();
pObj.addText('External link', { link: 'https://www.google.com' });
pObj.addLineBreak();
// Hyperlinks to bookmarks also supported: 
pObj.addText('Internal link', { hyperlink: 'myBookmark' });
// ... 
// Start somewhere a bookmark: 
pObj.startBookmark('myBookmark');
// ... 
// You MUST close your bookmark: 
pObj.endBookmark();

pObj.addLineBreak(); pObj.addLineBreak();


var pObj = docx.createListOfNumbers();

pObj.addText('Option 1');

var pObj = docx.createListOfNumbers();

pObj.addText('Option 2');

pObj.addLineBreak(); pObj.addLineBreak();

docx.putPageBreak();
var table = [
  [{
    val: "Sr. No.",
    opts: {
      cellColWidth: 4261,
      b: true,
      sz: '35',
      shd: {
        fill: "red",
        themeFill: "text1",
        "themeFillTint": "80"
      },
      fontFamily: "Avenir Book"
    }
  }, {
    val: "Title1",
    opts: {
      b: true,
      sz: '30',
      color: "yellow",
      align: "center",
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }, {
    val: "Title2",
    opts: {
      align: "right",
      vAlign: "center",
      cellColWidth: 42,
      b: true,
      sz: '48',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }],
  [1, 'text 1@', 'start'],
  [2, 'text 2$', ''],
  [3, 'text3?', ''],
  [4, 'text4!', 'END'],
]

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: "blue",
  tableAlign: "left",
  tableFontFamily: "Comic Sans MS",
  borders: true
}

docx.createTable(table, tableStyle);

var pObj = docx.createP();
pObj.addLineBreak(); pObj.addLineBreak();
pObj.addImage(path.resolve(IMAGEDIR, 'image (3).jpg'), { cx: 600, cy: 500 });


var FILENAME = "test-doc1.docx";
var out = fs.createWriteStream(OUTDIR + FILENAME);
docx.generate(out);

var docx = officegen('docx');
var pObj = docx.createP();

pObj.addText ( 'asd ',{font_size : 20 ,font_face:'Ravie'} );
pObj.addText ( 'Technology ', { back: '00ffff', shdType: 'pct12', shdColor: 'ff0000',font_size : 20 ,font_face:'Jokerman' } ); // Use pattern in the background.
pObj.addText ( ' And ' ,{font_size : 20 ,font_face:'Script MT'});
pObj.addText ( 'Services ', { highlight: true ,font_size : 20 ,font_face:'Small Fonts'} ); // Highlight!
pObj.addText ( 'asd. Ltd. asd', { highlight: 'darkGreen' ,font_size : 20 ,font_face:'Lucida Calligraphy'} ); // Different highlight color.

var pObj = docx.createListOfNumbers();

pObj.addText('Option 1');

var pObj = docx.createListOfNumbers();

pObj.addText('Option 2');

pObj.addLineBreak(); pObj.addLineBreak();

docx.putPageBreak();
var table = [
  [{
    val: "Sr. No.",
    opts: {
      cellColWidth: 4261,
      b: true,
      sz: '35',
      shd: {
        fill: "red",
        themeFill: "text1",
        "themeFillTint": "80"
      },
      fontFamily: "Avenir Book"
    }
  }, {
    val: "Title1",
    opts: {
      b: true,
      sz: '30',
      color: "yellow",
      align: "center",
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }, {
    val: "Title2",
    opts: {
      align: "right",
      vAlign: "center",
      cellColWidth: 42,
      b: true,
      sz: '48',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }],
  [1, 'text 1@', 'start'],
  [2, 'text 2$', ''],
  [3, 'text3?', ''],
  [4, 'text4!', 'END'],
]

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: "blue",
  tableAlign: "left",
  tableFontFamily: "Comic Sans MS",
  borders: true
}

docx.createTable(table, tableStyle);

var pObj = docx.createP();
pObj.addLineBreak(); pObj.addLineBreak();
pObj.addImage(path.resolve(IMAGEDIR, 'image (3).jpg'), { cx: 600, cy: 500 });


var FILENAME = "test-doc2.docx";
var out = fs.createWriteStream(OUTDIR + FILENAME);
docx.generate(out);


//------------------------------------------------------------------------------------------------------------

