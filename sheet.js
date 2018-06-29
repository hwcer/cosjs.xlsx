const RS = "\t\n";
const FS = "\t\r";
const XLSX  = require('xlsx');

function file_parse_xlsx(path,callback){
    let workbook = XLSX.readFile(path);
    workbook.SheetNames.forEach(function(sheetName) {
        let json = sheet_parse_json(workbook.Sheets[sheetName]);
        if(!json){
            console.log("Sheets Error",sheetName,path);
            json = {};
        }
        callback(json.name,json.rows);
    });
}


function sheet_parse_json(sheet){
    var sheet_data = XLSX.utils.sheet_to_csv(sheet,{"RS":RS,"FS":FS});

    var body = sheet_data.split(RS);
    if(body.length <4){
        return false;
    }
    var fname = body.shift().split(FS);
    var ftype = body.shift().split(FS);
    var fkeys = body.shift().split(FS);
    var ftext = body.shift();

    var fileName = fname[0];
    var fileType = fname[1]|| "json";
    var fileKeys = sheet_get_fields(fkeys);

    var data = {};
    for(var v of body){
        var val = sheet_fields_value(fileKeys,ftype,v.split(FS));
        var id = val["id"];
        if(!id){
            continue;
        }
        if(fileType =="json"){
            data[id] = val;
        }
        else if(fileType =="array"){
            if(!data[id]) {
                data[id]={"id":id,"name":val["name"]||"","rows":[]};
            }
            data[id]["rows"].push(val);
        }
        else if(fileType =="kv"){
            data[id] = val["val"]||"";
        }
    }
	return {"name":fileName,"type":fileType,"rows":data}
}

function sheet_fields_value(f,t,arr){
    var d = {};
    for(var k in f){
        d[k] = sheet_get_object(f[k],t,arr);
    }
    return d;
}


function sheet_get_object(i,t,arr){
    if(typeof i !="object"){
        return sheet_get_value(i,t,arr);
    }
    var v;
    if(Array.isArray(i)){
        v= [];
        for(var ii of i){
            v.push(sheet_get_object(ii,t,arr))
        }
    }
    else{
        v={};
        for(var ki in i){
            v[ki] =sheet_get_object(i[ki],t,arr)
        }
    }
    return v;
}


function sheet_get_value(i,t,arr){
    if(t[i]==='int' ){
        return parseInt(arr[i]||0)
    }
    else if(t[i]==='time' ){
        return arr[i] ? Date.parse(arr[i]) : 0;
    }
    else if(t[i]==='number' ){
        return parseFloat(arr[i]||0)
    }
    else{
        return String(arr[i]||'');
    }
}


function sheet_get_fields(arr){
    var fields = {},cache1=null,cache2=null,dep=0;
    for(var i=0;i<arr.length;i++){
        var k = arr[i];
        if(k.indexOf("[{") >=0){
            dep =2;
            cache1 = [],cache2={};
            var ak = k.split("[{");
            fields[ak[0]] = cache1;
            cache2[ak[1]] = i;
            cache1.push(cache2);
        }
        else if(k.indexOf("[[") >=0){
            dep =2;
            cache1 = [],cache2=[];
            var ak = k.split("[[");
            fields[ak[0]] = cache1;
            cache2.push(i)
            cache1.push(cache2);
        }
        else if(k.indexOf("[") >=0){
            if(!dep) dep =1;
            var ak = k.split("[");
            if(!cache1 || !cache2){
                cache1 = [];
                fields[ak[0]] = cache1;
            }
            if(cache2) cache2.push(i);
            else if (cache1) cache1.push(i);
        }
        else if(k.indexOf("{") >=0){
            if(!dep) dep =1;
            var ak = k.split("{");
            if(!cache1 || !cache2){
                cache1 = {};
                fields[ak[0]] = cache1;
            }
            var ak2 = ak[1];
            if(cache2) cache2[ak2] = i;
            else if (cache1) cache1[ak2] = i;
        }
        else if(k.indexOf("}]") >=0){
            dep =0;
            var ak = k.split("}]");
            cache2[ak[0]] = i;
            cache1=null,cache2=null;
        }
        else if(k.indexOf("]]") >=0){
            dep =0;
            var ak = k.split("]]");
            cache2.push(i);
            cache1=null,cache2=null;
        }
        else if(k.indexOf("]") >=0){
            var ak = k.split("]");
            if(dep>=2){
                cache2.push(i);
                cache2=[];
                cache1.push(cache2);
            }
            else{
                cache1.push(i);
                dep =0;
                cache1=null;
                cache2=null;
            }
        }
        else if(k.indexOf("}") >=0){
            var ak = k.split("}");
            if(dep==2){
                cache2[ak[0]] = i;
                cache2={};
                cache1.push(cache2);
            }
            else{
                cache1[ak[0]] = i;
                dep =0;
                cache1=null;
                cache2=null;
            }
        }
        else{
            if(dep==2) Array.isArray(cache2) ? cache2.push(i) :cache2[k] = i;
            else if(dep==1) Array.isArray(cache1) ? cache1.push(i) :cache1[k] = i;
            else {
                fields[k] = i;
            }
        }
    }
    return fields;
}


if(typeof exports !== 'undefined'){
    exports = module.exports = file_parse_xlsx;
}