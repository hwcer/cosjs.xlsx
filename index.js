
const fs = require("fs");
const path = require("path");



exports = module.exports = function (xlsx,json) {
    let sheet = require("./lib/sheet");
    let loader = require("./lib/loader")();
    loader.addPath(xlsx);
    loader.forEach((k,p)=>{
        let f = path.dirname(k);
        sheet(p,writeFile.bind(null,json,f));
    })
}



function writeFile(root,fname,sname,json){
    let file,dir = (fname && fname!=="/") ? fname : "";
    mkdir(root,dir);
    if(!dir){
        file = '/' + sname + ".json"
    }
    else{
        file = dir + '/' + sname + ".json"
    }

    fs.writeFile(root + file, JSON.stringify(json), (err) => {
        if (err) throw err;
        console.log("writeFile:",file);
    });
}


function mkdir(root,dir){
    if(!dir) return;
    let arr = dir.split("/");
    let p = [root];
    for(let v of arr){
        if(!v) continue;
        p.push(v);
        let f= p.join("/");
        if(!fs.existsSync(f)){
            fs.mkdirSync(f);
        }
    }
}