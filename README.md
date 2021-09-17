# How to use:

`npm i @lucafont2/onedrive-api`

```
var OnedriveApi = require("@lucafont2/onedrive-api")

(async function() {
     var onedrive = new OnedriveApi(YOUR_REFRESH_TOKEN,async function (error,userInfo) {
         if (error) {
             throw error;
         }  
     })
})();
```

The application provides you with a pre-built platform to get your refresh token
Go to https://lucaservers.com/onedrivetoken/ authorize and get the token.
In the future, you can revoke your token by going to https://account.live.com/consent/Manage

# Functions
Method | desc | params | response
--- | --- | --- | --- |
getDrives() | Gives you all drives in your account | none | array of drive objects |
getDrive(id) | Gives you all info of one drive| **id**: drive id got from getDrives() | info object |
getDirChildren(id,path) | Gives you all children folders and files in path | **id**: drive id, **path**: absolute path in the drive | array of files and folders info objects |
fileInfoByPath(id,filePath) | Gives file info by giving his path | **id**: drive id, **filePath** absolute path of file in the drive | file object |
fileInfoById(id, fileId) | Gives file info by giving his id  | **id**: drive id, **fileId** uniqId of the file got from getDirChildren() or fileInfoByPath() | file object |
downloadFileById(id,fileId,dir) | Download file by the id | **id**: drive id, **fileId**: uniqId of the file, **dir**: directory to place the download  | true if success |
deleteFileById(id,fileId) | Delete file by the id  | **id**: drive id, **fileId**: uniqId of the file | true if success |
moveFileById(id,fileId,newDir,newName) | Move a file in the drive | **id**: drive id, **fileId**: uniqId of the file, **newDir**: absolute path of the new folder | true if success |
copyFileById(id,fileId,newDir,newName) | Copy a file in the drive | **id**: drive id, **fileId**: uniqId of the file, **newDir**: absolute path of the new folder | true if success |
createDirectory(id,dir,name) | Create a dir in the drive | **id**: drive id, **dir**: directory to place the folder, **name**: name of the new dir | Info object of the new folder |
uploadFile(id,dir,name,localFile) | Upload a file | **id**: drive id, **dir**: directory to place the file, **name**: name of the file, **localFile**: string of the file path to upload | true if success |
getFileShareLink(id,fileId,props) | Get file share link | **id**: drive id, **fileId**: uniqId of the file, **props**: props object | link String |

#Objects

### Drive object
Key | Value | Description
--- | --- | --- |
id | String | Drive uniqId
type | String | personal / business / documentLibrary
owner | String | Name of owner

### File or folder object
Key | Value | Description
--- | --- | --- |
id | String | Element uniqId
name | Sting | Element name
size | Int | Element bytes
lastModified | String | Last Modified
parent | String | Parent path
type | String | File / Folder
typeInfo | Object | if type is File contains mime and sha1

### Props object
Key | Value | Description
--- | --- | --- |
type | String | Permission type view/edit/embed
password | OPTIONAL - String | Password protected link
scope | OPTIONAL - String | anonymous / organization
expirationDateTime | OPTIONAL - String | Link exipiration

# Functions
All functions work with async - promise

```
var OnedriveApi = require("@lucafont2/onedrive-api")

(async function() {
 var onedrive = new OnedriveApi(token,async function (error,userInfo) {
     try {
         if (error) {
             throw error;
         }
         console.log("Logged in with", userInfo.displayName);
         var drives = await onedrive.getDrives();
         console.log(drives);
         var drive = await onedrive.getDrive(drives[0]["id"]);
         console.log(drive);
     }catch (e) {
         console.log(e)
     }
 })

})();
```
