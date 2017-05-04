var demo = require("npm-demo-pkg");

demo.printMsg("red", "trying to create files...");
demo.printMsg("yellow", "------------------------------");

var path = require('./doc_ex');
demo.printMsg("cyan", "|- test-doc1.docx");
demo.printMsg("cyan", "|- test-doc2.docx");

var path = require('./excel_ex');
demo.printMsg("cyan", "|- test-excel1.xlsx");
demo.printMsg("yellow", "   |-- Sheet1");
demo.printMsg("yellow", "   |-- Sheet2");

demo.printMsg("yellow", "------------------------------");
demo.printMsg("green", "File created successfully");