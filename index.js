class OnedriveApi{
    axios = require("axios");
    fs=require("fs");
    path=require("path");

    auth={};
    info={};
    downloadUrl = {};
    constructor(refresh_token,callback) {
        this.auth.refresh_token = refresh_token;

        this.auth.constructed = true;
        this.login(callback);
    }
    login(callback){
        var oneThis = this;
        if (this.auth.constructed){
            this.axios({
                "url": "https://lucaservers.com/onedrivetoken/refreshToken.php?refresh="+oneThis.auth.refresh_token,
            }).then((r)=>{
                if (!r.data.hasOwnProperty("token") || r.data.token === null){
                    callback({"error":"Invalid refresh token"});
                    return;
                }
                var token = r.data.token;

                if ( token.length>10){
                    oneThis.auth.onedrive_token = token;
                    oneThis.axios({
                        "url": "https://graph.microsoft.com/v1.0/me/",
                        "headers":{
                            "Authorization":"Bearer " + oneThis.auth.onedrive_token
                        }
                    }).then((r)=>{

                        oneThis.auth.userInfo = r.data;
                        oneThis.auth.success = true;

                        callback(null,r.data);
                    }).catch((e)=>{

                        oneThis.auth.success = false;
                        callback(null,e.response.data);
                    })
                }else{
                    callback({"error":"Invalid refresh token"});
                }
            }).catch((e)=>{
                console.log(e);
                oneThis.auth.success = false;
                callback(e.response.data);
            })

        }
    }
    async getDrives(){
        var oneThis = this;


        return new Promise(function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            oneThis.axios({
                "url": "https://graph.microsoft.com/v1.0/me/drives",
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                }
            }).then((r)=>{

                oneThis.info.drives = r.data;
                var parsed = [];
                for (var i = 0;i<r.data.value.length;i++){
                    let a = r.data.value[i];
                    parsed.push({"id":a.id,"type":a.driveType,"owner":a.owner.user.displayName});
                }
                resolve(parsed);


            }).catch((e)=>{
                reject({"error":e.response.data});

            })
        })
    }
    async getDrive(driveId){
        var oneThis = this;


        return new Promise(function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/root`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                }
            }).then((r)=>{
                resolve(r.data);
            }).catch((e)=>{
                reject({"error":e.response.data});
            })
        })
    }
    async getDirChildren(driveId,path="/"){
        var oneThis = this;

        if (path!=="/"){
            path = encodeURIComponent(path);

            path = ":" + path + ":";
        }
        return new Promise(function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/root${path}/children`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                }
            }).then((r)=>{
                var parsed = [];
                for (var i = 0;i<r.data.value.length;i++){
                    let a = r.data.value[i];
                    var type = "";
                    var info = {};
                    if (a.hasOwnProperty("folder")){
                        type="folder";
                        info = {};
                    }else if(a.hasOwnProperty("file")){
                        type = "file";
                        info = {
                            "hash":a.file.hashes,
                            "mime":a.file.mimeType,
                            "download":a.id
                        };
                        oneThis.downloadUrl[a.id]= {"u":a["@microsoft.graph.downloadUrl"],name:a.name};
                    }

                    parsed.push({"id":a.id,"name":a.name,"size":a.size,"lastModified":a.lastModifiedDateTime,"parent":a.parentReference.path.replace(":",""),type:type,typeInfo:info});
                }
                resolve(parsed);
            }).catch((e)=>{
                reject({"error":e.response.data});
            })
        })
    }
    async doDownload(url,path){
        var oneThis = this;

        return new Promise(function (resolve, reject) {
            if (oneThis.fs.existsSync(path)){
                oneThis.fs.unlinkSync(path);
            }
            oneThis.axios({
                "url": url,
                responseType: 'stream'

            }).then((r) => {
                const writer = oneThis.fs.createWriteStream(path);
                r.data.pipe(writer)
                writer.on('finish', () => {
                    resolve(true);
                })
                writer.on('error', (er) => {
                    reject({"error": er});
                })


            }).catch((e) => {
                reject({"error": e.response.data});

            })
        })
    }
    async downloadFileById(driveId,fileId,path="./"){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            if (!oneThis.downloadUrl.hasOwnProperty(fileId)){
                oneThis.axios({
                    "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}`,
                    "headers":{
                        "Authorization":"Bearer " + oneThis.auth.onedrive_token
                    }
                }).then(async function (r){
                    var pathF = oneThis.path.join(path,r.data.name);
                    resolve(await oneThis.doDownload(r.data["@microsoft.graph.downloadUrl"],pathF))
                }).catch((e)=>{
                    reject({"error":e.response.data});
                })
            }else{
                var pathF = oneThis.path.join(path,oneThis.downloadUrl[fileId]["name"]);

                resolve(await oneThis.doDownload(oneThis.downloadUrl[fileId]["u"],pathF))

            }

        })
    }
    async fileInfoById(driveId,fileId){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            if (!oneThis.downloadUrl.hasOwnProperty(fileId)){
                oneThis.axios({
                    "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}`,
                    "headers":{
                        "Authorization":"Bearer " + oneThis.auth.onedrive_token
                    }
                }).then(async function (r){
                    let a = r.data;
                    var type = "";
                    var info = {};
                    if (a.hasOwnProperty("folder")){
                        type="folder";
                        info = {};
                    }else if(a.hasOwnProperty("file")){
                        type = "file";
                        info = {
                            "hash":a.file.hashes,
                            "mime":a.file.mimeType,
                            "download":a.id
                        };
                        oneThis.downloadUrl[a.id]= {"u":a["@microsoft.graph.downloadUrl"],name:a.name};
                    }

                    resolve({"id":a.id,"name":a.name,"size":a.size,"lastModified":a.lastModifiedDateTime,"parent":a.parentReference.path.replace(":",""),type:type,typeInfo:info});

                }).catch((e)=>{
                    reject({"error":e.response.data});
                })
            }else{
                resolve(oneThis.downloadUrl[fileId])
            }

        })
    }
    async fileInfoByPath(driveId,filePath){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            filePath = encodeURIComponent(filePath);

            filePath = ":" + filePath + ":";
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/root${filePath}/`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                }
            }).then(async function (r){
                let a = r.data;
                var type = "";
                var info = {};
                if (a.hasOwnProperty("folder")){
                    type="folder";
                    info = {};
                }else if(a.hasOwnProperty("file")){
                    type = "file";
                    info = {
                        "hash":a.file.hashes,
                        "mime":a.file.mimeType,
                        "download":a.id
                    };
                    oneThis.downloadUrl[a.id]= {"u":a["@microsoft.graph.downloadUrl"],name:a.name};
                }

                resolve({"id":a.id,"name":a.name,"size":a.size,"lastModified":a.lastModifiedDateTime,"parent":a.parentReference.path.replace(":",""),type:type,typeInfo:info});

            }).catch((e)=>{
                reject({"error":e.response.data});
            })

        })
    }
    async deleteFileById(driveId,fileId){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }

            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                },
                "method":"DELETE"
            }).then(async function (r){
                resolve(true)

            }).catch((e)=>{
                reject({"error":e.response.data});
            })


        })
    }
    async moveFileById(driveId,fileId,parentNewFolderId,newName){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            var body = {
                "parentReference": {
                    "id": parentNewFolderId
                },
                "name": newName
            }
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                },
                "method":"PATCH",
                data:body
            }).then(async function (r){
                resolve(true)

            }).catch((e)=>{
                reject({"error":e.response.data});
            })


        })
    }
    async copyFileById(driveId,fileId,parentNewFolderId,newName){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            var body = {
                "parentReference": {
                    "id": parentNewFolderId
                },
                "name": newName
            }
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}/copy`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                },
                "method":"POST",
                data:body
            }).then(async function (r){
                resolve(true)

            }).catch((e)=>{
                reject({"error":e.response.data});
            })


        })
    }
    async createDirectory(driveId,drivePath,folderName){
        var oneThis = this;

        if (drivePath!=="/"){
            drivePath = encodeURIComponent(drivePath);

            drivePath = ":" + drivePath + ":";
        }
        return new Promise(function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            var body =
            {
                "name": folderName,
                "folder": {}
            }

            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/root${drivePath}/children`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                },
                method:"POST",
                data:body
            }).then((r)=>{
                let a = r.data;
                var type = "";
                var info = {};
                if (a.hasOwnProperty("folder")){
                    type="folder";
                    info = {};
                }else if(a.hasOwnProperty("file")){
                    type = "file";
                    info = {
                        "hash":a.file.hashes,
                        "mime":a.file.mimeType,
                        "download":a.id
                    };
                    oneThis.downloadUrl[a.id]= {"u":a["@microsoft.graph.downloadUrl"],name:a.name};
                }

                resolve({"id":a.id,"name":a.name,"size":a.size,"lastModified":a.lastModifiedDateTime,"parent":a.parentReference.path.replace(":",""),type:type,typeInfo:info});

            }).catch((e)=>{
                reject({"error":e.response.data});
            })
        })
    }
    //https://gist.github.com/tanaikech/22bfb05e61f0afb8beed29dd668bdce9
    getparams(file){
        var allsize = this.fs.statSync(file).size;
        var sep = allsize < (60 * 1024 * 1024) ? allsize : (60 * 1024 * 1024) - 1;
        var ar = [];
        for (var i = 0; i < allsize; i += sep) {
            var bstart = i;
            var bend = i + sep - 1 < allsize ? i + sep - 1 : allsize - 1;
            var cr = 'bytes ' + bstart + '-' + bend + '/' + allsize;
            var clen = bend != allsize - 1 ? sep : allsize - i;
            var stime = allsize < (60 * 1024 * 1024) ? 5000 : 10000;
            ar.push({
                bstart : bstart,
                bend : bend,
                cr : cr,
                clen : clen,
                stime: stime,
            });
        }
        return ar;
    }
    async uploadFile(driveId,folder,fileName,localFile) {
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.fs.existsSync(localFile)){
                reject({"error":"File not exists"})
                return;
            }
            try{
                var upUrl = async function(){
                    if (folder.slice(-1)!=="/"){
                        folder+="/";
                    }
                    folder+=fileName;
                    folder = encodeURIComponent(folder);

                    folder = ":" + folder + ":";

                    return new Promise(async function(resolve1,reject1){
                        var body = {
                            "@microsoft.graph.conflictBehavior": "rename",
                            "name": fileName,
                            "fileSize": oneThis.fs.statSync(localFile).size,
                        };
                        oneThis.axios({
                            "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/root${folder}/createUploadSession`,
                            "headers":{
                                "Authorization":"Bearer " + oneThis.auth.onedrive_token
                            },
                            method:"POST",
                            data:body
                        }).then((r)=>{
                            resolve1(r.data.uploadUrl);
                        }).catch(( e)=>{
                            console.log(e.response.data)
                            reject1({"error":e.response.data});
                        })
                    });

                }
                var fileData = oneThis.fs.readFileSync(localFile);
                var uploadChunk = async function(elem,urlU){
                    return new Promise(async function(resolve2,reject2){
                        oneThis.axios({
                            "url": urlU,
                            method:"PUT",
                            headers: {
                                'Content-Length': elem.clen,
                                'Content-Range': elem.cr,
                            },
                            data:fileData.slice(elem.bstart, elem.bend + 1),
                            maxContentLength: (70 * 1024 * 1024),
                            maxBodyLength: (70 * 1024 * 1024)
                        }).then((r)=>{
                            r=null;
                            resolve2(true);
                        }).catch(( e)=>{
                            console.log(e.toJSON());
                            reject2({"error":e.response.data});
                            e=null;
                        })
                    })
                }
                var url = await upUrl();
                var params = oneThis.getparams(localFile);

                for (var i  = 0;i<params.length;i++){
                    await uploadChunk(params[i],url);
                }
                resolve(true);
            }catch (e) {
                reject(e);
            }


        });
    }
    async getFileShareLink(driveId,fileId,props){
        var oneThis = this;
        return new Promise(async function (resolve, reject) {
            if (!oneThis.auth.success){
                reject({"error":"not logged in"});
                return;
            }
            var options = {"type":["view","edit","embed"],"scope":["anonymous","organization"]};
            var body = {};
            if (props.hasOwnProperty("type")){
                if (options.type.includes(props.type)){
                    body.type = props.type;
                }else{
                    reject({"error":"Not valid type"})
                    return;
                }
            }else{
                reject({"error":"Must enter type"});
                return ;

            }
            if (props.hasOwnProperty("scope")){
                if (options.scope.includes(props.scope)){
                    body.scope = props.scope;
                }else{
                    reject({"error":"Not valid scope"});
                    return ;
                }
            }
            if (props.hasOwnProperty("expirationDateTime")){
                body.expirationDateTime=props.expirationDateTime;
            }
            if (props.hasOwnProperty("password")){
                body.password=props.password;
            }
            oneThis.axios({
                "url": `https://graph.microsoft.com/v1.0/me/drives/${driveId}/items/${fileId}/createLink`,
                "headers":{
                    "Authorization":"Bearer " + oneThis.auth.onedrive_token
                },
                "method":"POST",
                data:body
            }).then(async function (r){
                resolve(r.data.link)

            }).catch((e)=>{
                reject({"error":e.response.data});
            })


        })
    }
}
module.exports = OnedriveApi;

