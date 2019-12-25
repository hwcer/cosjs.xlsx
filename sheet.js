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

if(typeof exports !== 'undefined'){
    exports = module.exports = file_parse_xlsx;
}

function sheet_parse_json(sheet){

    //let sheet_data = XLSX.utils.sheet_to_csv(sheet,{"RS":RS,"FS":FS});
    let sheet_data =  XLSX.utils.sheet_to_json(sheet, {header:1});
    if(sheet_data.length <4){
        return false;
    }
    let fname = sheet_data.shift();
    let ftype = sheet_data.shift();
    let fkeys = sheet_data.shift();
    let ftext = sheet_data.shift();

    let fileName = fname[0];
    let fileType = fname[1]|| "json";
    let fileKeys = sheet_get_fields(fkeys);

    let data = {};
    for(let v of sheet_data){
        if(!v){
            continue;
        }
        let val = sheet_fields_value(fileKeys,ftype,v);
        if( !val || !("id" in val)){
            continue;
        }
        let id = val["id"];
        if(fileType =="json"){
            data[id] = val;
        }
        else if(fileType =="array"){
            if(!data[id]) {
                data[id]=[];
            }
            data[id].push(val);
        }
        else if(fileType =="kv"){
            data[id] = val["val"]||"";
        }
    }
	return {"name":fileName,"type":fileType,"rows":data}
}

function sheet_fields_value(f,t,arr){
    //检查是不是空行
    if(arr.length === 0 || (!arr[0] && typeof arr[0] !== 'number' ) ){
        return false;
    }
    let d = {};
    for(let k in f){
        d[k] = sheet_get_object(f[k],t,arr);
    }
    return d;
}


function sheet_get_object(i,t,arr){
    if(typeof i !="object"){
        return sheet_get_value(i,t,arr);
    }
    let v;
    if(Array.isArray(i)){
        v= [];
        for(var ii of i){
            v.push(sheet_get_object(ii,t,arr))
        }
    }
    else{
        v={};
        for(let ki in i){
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
    let fields = {},cache1=null,cache2=null,dep=0;
    for(let i=0;i<arr.length;i++){
        let k = arr[i];
        if(!k){
            continue;
        }
        k = k.replace(/<[^>]+>/g,"");
        if(k.indexOf("[{") >=0){
            dep =2;
            cache1 = [],cache2={};
            let ak = k.split("[{");
            fields[ak[0]] = cache1;
            cache2[ak[1]] = i;
            cache1.push(cache2);
        }
        else if(k.indexOf("[[") >=0){
            dep =2;
            cache1 = [],cache2=[];
            let ak = k.split("[[");
            fields[ak[0]] = cache1;
            cache2.push(i)
            cache1.push(cache2);
        }
        else if(k.indexOf("[") >=0){
            if(!dep) dep =1;
            let ak = k.split("[");
            if(!cache1 || !cache2){
                cache1 = [];
                fields[ak[0]] = cache1;
            }
            if(cache2) cache2.push(i);
            else if (cache1) cache1.push(i);
        }
        else if(k.indexOf("{") >=0){
            if(!dep) dep =1;
            let ak = k.split("{");
            if(!cache1 || !cache2){
                cache1 = {};
                fields[ak[0]] = cache1;
            }
            let ak2 = ak[1];
            if(cache2) cache2[ak2] = i;
            else if (cache1) cache1[ak2] = i;
        }
        else if(k.indexOf("}]") >=0){
            dep =0;
            let ak = k.split("}]");
            cache2[ak[0]] = i;
            cache1=null,cache2=null;
        }
        else if(k.indexOf("]]") >=0){
            dep =0;
            let ak = k.split("]]");
            cache2.push(i);
            cache1=null,cache2=null;
        }
        else if(k.indexOf("]") >=0){
            let ak = k.split("]");
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
            let ak = k.split("}");
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


