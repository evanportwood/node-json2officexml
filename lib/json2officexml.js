var xmlbuilder = require('xmlbuilder');
    
var j2o = exports;

var XMLBANNEDCHARS = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;


var ExcelOfficeXmlWriter = j2o.ExcelOfficeXmlWriter = function(options) {

};

ExcelOfficeXmlWriter.prototype.writeDoc = function(doc) {
    if (!doc) return;
    return _writeExcelDoc(this, doc);
};

j2o.createExcelOfficeXmlWriter = function(path, options) {
    return new ExcelOfficeXmlWriter(options);
};

function _isoDateString(d){  
    function pad(n){return n<10 ? '0'+n : n}  
    return d.getUTCFullYear()+'-'
    + pad(d.getUTCMonth()+1)+'-'  
    + pad(d.getUTCDate())+'T'  
    + pad(d.getUTCHours())+':'  
    + pad(d.getUTCMinutes())+':'  
    + pad(d.getUTCSeconds())+".000"  ;//'Z'  
}

function _writeExcelDoc(writer, o) {
    var XMLHDR = { 'version': '1.0'};
    var doc = xmlbuilder.create();
    var child = doc.begin('Workbook', XMLHDR);
    //add workbook attributes
    child.att("xmlns","urn:schemas-microsoft-com:office:spreadsheet");
    child.att("xmlns:o","urn:schemas-microsoft-com:office:office");
    child.att("xmlns:x","urn:schemas-microsoft-com:office:excel");
    child.att("xmlns:ss","urn:schemas-microsoft-com:office:spreadsheet");
    child.att("xmlns:html","http://www.w3.org/TR/REC-html40");
    //add document properties and attributes
    child = child.ele("DocumentProperties").att("xmlns","urn:schemas-microsoft-com:office:office");
    child = child.ele("Created").raw( _isoDateString(new Date())).up().up();
    //add header row
    if (o.columns.length) {
        //set child to workbook and add header row styles
        var worksheet = child;
        child = child.root();
        child.ele("ss:Styles").ele("ss:Style").att("ss:ID", "1").ele("ss:Font").att("ss:Bold", "1").up().up().up();
        //add worksheet
        child = child.ele("ss:Worksheet").att("ss:Name", "Sheet1").ele("ss:Table");
        //add header row cells
        child = child.ele("ss:Row").att("ss:StyleID", "1");
        o.columns.forEach(function (v,i) {
            //only allow strings as row header values
            if (typeof v.header === 'string') {
                var str = v.header.split('\u000b').join(' ');
                child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(str).up().up();
            }
        });
        child = child.up();
        //add content rows that are column specific
        o.rows.forEach(function (val, i) {
            child = child.ele("ss:Row");
            o.columns.forEach(function (v, i) {
                if (val[v.data]) {
                    //column value exists in row
                    if (typeof val[v.data] !== 'function') {
                        if (typeof val[v.data] === 'object') {
                            if (val[v.data] instanceof Date) {
                                child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "DateTime").raw(_isoDateString(new Date(val[v.data]))).up().up();
                            } else {
                                if (val[v.data] instanceof Array) { }
                            }
                        } else {
                            if ((typeof val[v.data]) === 'boolean') {
                            } else if ((typeof val[v.data]) === 'number') {
                                child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "Number").txt(val[v.data]).up().up();
                            } else {
                                var str = val[v.data].split('\u000b').join(' ');
                                child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(str).up().up();
                            }
                        }
                    }
                } else {
                    //column value does not exist in row - add empty cell
                    child = child.ele("ss:Cell").up();
                }
            });

            child = child.up();
        });
    } else {
        //add worksheet
        child = child.ele("ss:Worksheet").att("ss:Name", "Sheet1"); //.ele("ss:Table");
        //add content rows that are column agnostic
        o.rows.forEach(function (i, v) {
            child = child.ele("ss:Row");
            for (name in i) {
                if (typeof i[name]!== 'function') {
                    if (typeof i[name]=== 'object') {
                        if (i[name] instanceof Date) {
                            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "DateTime").raw(_isoDateString(new Date(i[name]))).up().up();
                        } else {
                            if (i[name] instanceof Array) { }
                        }
                    } else {
                        if ((typeof i[name]) === 'boolean') {
                        } else if ((typeof i[name]) === 'number') {
                            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "Number").txt(i[name]).up().up();
                        } else {
                                    //chr = str.match(chars);
                            var str = i[name].split('\u000b').join(' ');
                            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(str).up().up();
                        }
                    }
                }
            }
            child = child.up();
        });
        //set child back to worksheet
        child = child.up();
    }

    /*
    //add worksheet options
    child = child.ele("WorksheetOptions").att("xmlns","urn:schemas-microsoft-com:office:excel");
    child = child.ele("ProtectObjects").txt("True").up();
    child = child.ele("ProtectScenarios").txt("True").up();

    //set child back to worksheet
    child = child.up();
    */

    return doc;
}