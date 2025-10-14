
// graphFileUtility.js

  'use strict'
  console.log('Loading Module graphFileUtil.js') // This is useful to trace any issues with load sequence
  const semVer = {major:2, minor:1, build: 0}
  


  // Define Internal variables here
  const _getOptions = { headers: { Accept: 'application/json; odata=nometadata' } }
  

  // Define various method functions as required
  export async function getSiteDriveFileHistory (site, drive, file, filters, limitCount = 0) {
    if (!site || !drive || !file) {
      console.error(`Invalid data used to get get file history!\nSite: ${site}\nDrive: ${drive}\nFile: ${file}`)
      return []
    }
    try {

      if (typeof site === 'string') {
        if (!site.endsWith('/')) site += '/'
      }
      let url = `${site}_api/v2.1/drives/${drive}/items/${file}/versions`
      let urlParamSepChar = '?'
      if (typeof filters === 'string' && filters.length > 0) {
        url +=  urlParamSepChar + 'filter=' + filters
        urlParamSepChar = '&'
      }
      if (typeof limitCount === 'number' && limitCount >= 1) {
        url += urlParamSepChar + 'top=' + limitCount.toFixed(0)
      }

      // console.log(`Versions REST from\n${url}`)
      let resp = await fetch (url, _getOptions)
      let data = await resp.json()
      return data.value
    } catch (ex) {
      console.error(`Unexpected error in attempt to read file history!\nSite: ${site}\nDrive: ${drive}\nFile: ${file}\nError: ${ex instanceof Error ?  ex.message : ex?.toString()}`)
      return []
    }
  }
  
  export async function getSiteDriveFileDetails (site, drive, file) {
    if (typeof site === 'string') {
      if (!site.endsWith('/')) site += '/'
    }
    let url = `${site}_api/v2.1/drives/${drive}/items/${file}`
    // console.log(`Versions REST from\n${url}`)
    let resp = await fetch (url, _getOptions)
    let data = await resp.json()
    return data
  }
  export async function getSiteDrives (siteUrl) {
    let url =`${siteUrl}/_api/v2.1/drives`
    try {
    let resp = await fetch(url,_getOptions)
    let data = await resp.json()
    return data.value
    } catch (ex) {
      console.error(`Failed in graphFileUtility::getSiteDrives\nURL:\n${url}\n${ex.message}`)
      throw ex
    }
  }

  export async function getSiteDrivePathFolderContent (siteUrl, driveId, path, filesFoldersOrBoth = 'BOTH') {
    let sourceFolderId = '' // Will be populate after an initial API call to verify that the path is valid and a folder
    try {
      let urlToGetEntryFolder =`${siteUrl}/_api/v2.1/drives/${driveId}/root:${path}` // IMPORTANT: must ensure that path has trailing slash to get the children endpoint!
      let respFolderQuery = await fetch(urlToGetEntryFolder,_getOptions)
      let dataFolderQuery = await respFolderQuery.json()

      if (dataFolderQuery.hasOwnProperty('folder')) {
        sourceFolderId = dataFolderQuery.id
      } else {
        throw new Error(`The requested path (${path}) does not result in a folder.`)
      }
  
    } catch (error) {
      let msg = error instanceof Error ? error.message : error?.toString()
      throw new Error(msg)
    }
    
    let urlToGetChildren = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${sourceFolderId}/children`
    let filterOption = filesFoldersOrBoth.toUpperCase()
    switch(filterOption) {
      case 'FILES':
      case 'FILE':
        urlToGetChildren += '?filter=folder eq null'
        break
      case 'FOLDERS':
      case 'FOLDER':
        urlToGetChildren += '?filter=file eq null'
        break
      default: 
        break
    }
    let respChildrenQuery = await fetch(urlToGetChildren, _getOptions)
    let dataChildrenQuery = await respChildrenQuery.json()
    return dataChildrenQuery.value
  }

  export async function getSiteDriveFolderIdChildren (siteUrl, driveId, folderId) {
    let urlToGetChildren = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${folderId}/children`
    let respChildrenQuery = await fetch(urlToGetChildren, _getOptions)
    let dataChildrenQuery = await respChildrenQuery.json()
    return dataChildrenQuery.value
  }

  export async function getGraphItemFolderPath(item){
    if (item?.parentReference.id) {
      let pRef = item.parentReference
      let folder = await getSiteDriveFileDetails(this.siteUrl, pRef.driveId, pRef.id)
      if (folder?.webUrl){
        let siteUrl = pRef.sharepointIds.siteUrl
        
        let idxPathStart = siteUrl.indexOf('.com') + 4 + 1 // extra 1 character to move past the slash
        let path =   folder.webUrl.substring(idxPathStart)
        return path
      }
    }
  }

  function getVersion(){
    return `${semVer.major}.${semVer.minor}.${semVer.build}`
  }


export default  graphFileUtility = {
    getSiteDriveFolderIdChildren,
    getSiteDrivePathFolderContent,
    getSiteDrives,
    getSiteDriveFileHistory,
    getSiteDriveFileDetails,
    getGraphItemFolderPath,
    getVersion
}