"use strict";
var fs = require('fs');
//模块加载器


function loader(){
    if (!(this instanceof loader)) {
        return new loader();
    }
    this.ext = ['.xls','.xlsx'];
    this._pathCache   = new Set();
    this._moduleCache = {};
}

module.exports = loader;


loader.prototype.addPath = function(path) {
    if(!path){
        return;
    }
    if(this._pathCache.has(path)){
        return;
    }
    this._pathCache.add(path);
    getFiles.call(this, path);
}

loader.prototype.forEach = function(callback){
    for(let k in this._moduleCache){
        callback(k,this._moduleCache[k]);
    }
}

function getFiles(root,dir) {
    dir = dir||'/';
    var self = this,stats;
    var path = root + dir;
    try {
        stats = fs.statSync(path);
    }
    catch (e){
        stats = null;
    }
    if(!stats){
        return;
    }
    if(!stats.isDirectory()){
        return;
    }
    var files = fs.readdirSync(path);
    if(!files || !files.length){
        return;
    }
    files.forEach(function (name) {
        filesForEach.call(self,root,dir,name);
    })
}


function filesForEach(root,dir,name){
    var self = this,FSPath = require('path');
    var realPath = FSPath.resolve([root,dir,name].join('/'));
    var realName = dir+name;
    var stats = fs.statSync(realPath);
    if (stats.isDirectory()) {
        return getFiles.call(self,root,realName+'/');
    }

    var ext = FSPath.extname(name);
    if(self.ext.indexOf(ext) >=0){
        var api = realName.replace(ext,'');
        if( self._safeMode && self._moduleCache[api] ){
            console.log('file['+api+'] exist');
        }
        else{
            self._moduleCache[api] = realPath;
        }
    }
}