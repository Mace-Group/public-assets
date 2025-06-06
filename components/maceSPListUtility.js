/**
 * This script creates an old fashioned "module" (or namespace) pattern object
 * It provides utility functions to help CRUD operations in SharePoint lists
 * 
 * Designers note
 * That being said it DOES have a default base URL, which is detected from the *site collection* address. 
 * The caveat is that the methods exported here are not PURE because they rely updo data not specifically 
 * passed in the call signature, i.e. the Site_Url value is not passed with every call, although it ca be overridden if required!
 *  
 * General recommendations
 * Use "this" module wrapped by an additional application specific library which  in turn initialises the SITE_URL
 * and exposes appropriate lists by partial application of these methods
 * Example:-
 * Create a library for Widgets called widgetUtility
 * implement the widgetUtility with methods such as 
 * widgetUtility.getWidgets(options) which is to be defined as:-
 * namespace.getWidgets = maceSPListUtility.getListItems.bind(this, 'widget-list-name-or-guid', [['Title', 'widgetName', 'Widget Name'],['field_1', 'code', 'Widget #'], ...])
 * 
 * This will result in application code being able to use widget library like this:-
 * let aWidgets = await widgetUtility.getWidgets()
 * 
 * WARNING: 
 * The above example ignores the options parameter. If this is to be used you MUST use teh SharePoint field names and not logical property names
 * e.g.
 * 
 * let aWidgets = await widgetUtility.getWidgets({filerClause: "code eq 'XYZ' "}) ==> ERROR there is no field called "code"
 * let aWidgets = await widgetUtility.getWidgets({filerClause: `${maceSPListUtility.getLogicalToPhysicalFieldNameForMap(mapFieldsArray, 'code')} eq 'XYZ' `}) ==> SUCCESS would translate "code" to "field_1" for example
 * 
 * The above could be improved by wrapping the logical to physical like this
 * namespace.lookupWidgetField = maceSPListUtility.getLogicalToPhysicalFieldNameForMap.bind(this, mapFieldsArray)
 * 
 * so that the code to get filtered widgets becomes
 * let aWidgets = await widgetUtility.getWidgets({filerClause: `${widgetUtility.lookupWidgetField('code')} eq 'XYZ' `}) ==> SUCCESS would translate "code" to "field_1" for example
 * 
 * Darren H. Gill
 */


// ~~CONFIGURE HERE~~ set the global name required
globalThis.maceSPListUtility = (function (namespaceObject) {
  'use strict'
  console.log('Loading Module maceSPListUtility.js') // This is useful to trace any issues with load sequence
  const SemVer='2.1.1'
  /**
   * The SharePoint lists managed by this module were mainly created by the SharePoint List from Spreadsheet facility
   * The result is that many of the internal field names are unsuited to logical coding conventions (i.e field_1, field_2 etc)
   * Consequently this library has a translation capability used to map between the obscure internals return from the REST API to
   * useful and consistent names for the consumers of THIS library.
   */
  let SITE_BASE // = 'https://mace365.sharepoint.com/sites/MoJControlCentre' // ~~CONFIGURE HERE~~
  if (document?.currentScript) {
    let src = document.currentScript.src
    let reFindSite = /(.+)\/siteAssets\/(.*)/i
    if (reFindSite.test(src)){
      let match = reFindSite.exec(src)
      console.log(`maceSPListUtility dynamically determined its site base to be ${match[1]}`)
      SITE_BASE = match[1]
    }
  }
  
  let cacheCurrentUserInSite = {}
  const OK = 'OK'
  const FAIL = 'FAIL'
  const cacheFailedRequests = []
  function getSiteBase () { return SITE_BASE}
  function setSiteBase (url) { SITE_BASE = url}
  
  // Upscale Excel to List Mapping conversions
  // The following maps convert SharePoint List internal IDs to object properties to make coding life easier!
  
  const MAP_P2L_COMMON = [ // It will always be useful to know the Id, when last Modified (and who by)
    ['Modified', 'Modified', 'Modified Date'],
    ['Id','Id', 'Internal ID', 'SharePoint List Id'],
    ['odata.etag', 'eTag', 'Edit Tag'],
    ['odata.id', 'oDataId', 'OData Identifier']
  ] 

  // Help problem diagnosis with a list of dependencies that THIS module will use
  const assumedLoadedGlobals = [
    {globalName: 'R', description: 'Ramda 0.29'},
    {globalName: 'moment', description: 'Moment 2.29'}
  ]
  assumedLoadedGlobals.forEach(m => {
    if (!globalThis[m.globalName]) {
      console.warn(`Assumed global "${m.globalName}" for ${m.description} is not loaded. Errors will be likely.`)
    }
  })




  // Utility functions 
  const mapPickFirst = R.view(R.lensIndex(0))  // picks 1st column from 2D grid array
  const mapPickSecond = R.view(R.lensIndex(1)) // picks 2nd column from 2D grid array
  const mapPickThird = R.view(R.lensIndex(2)) // picks 3rd column from 2D grid array
  const mapPickFourth = R.view(R.lensIndex(3)) // picks 4thd column from 2D grid array
  const fnGetInternalNamesByTypeFromMap = (typeName, mapArray) => mapArray.reduce((acc,cur)=> {if(cur[3] === typeName) acc.push(cur[0]); return acc}, [])
  const fGetItemArrayFromRestData = R.pathOr([],['value']) // An implicitly curried fn call just needs a REST/JSON response
  const reIsISODateText = /(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(\.\d{1,})?)([Z+-]\d{0,4})/i
  const reIsGuidText =/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
  const localDateToSharePoint = function (value) { return moment(value).toISOString().substring(0,19) + 'Z'  }

  
function autoPrefixSiteUrlWithBase (siteBase) {
  if (typeof siteBase === 'string' && !siteBase.startsWith('https://')) {
    return `${SITE_BASE}${ siteBase.startsWith('/') ? '' : '/'}${siteBase}` //prefix the argument with <standard Site Url>/
  } else {
    return siteBase // They passed something weird so give it back!
  }
}

/**
 * @description internal translate the name to main part of rest URL
 * @param {string} listName 
 * @param {string} siteBase - the URL to the SharePoint site when it is NOT the standard as configured in this library (default to standard if ommitted)
 * @returns 
 */
  const getSiteListRestUrlByName = function (listName, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    return `${siteBase}/_api/web/lists/GetByTitle('${listName}')`
  }

/**
 * @description internal translate the name to main part of rest URL
 * @param {string} listGuid
 * @param {string} siteBase - the URL to the SharePoint site when it is NOT the standard as configured in this library (default to standard if ommitted)
 * @returns 
 */
  const getSiteListRestUrlByGuid = function (guid, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    return `${siteBase}/_api/web/lists(guid'${guid}')`
  }

  const fConvertIsoDateTextToDate = function (dateText) {
    if (reIsISODateText.test(dateText)){
      
      let mCol = reIsISODateText.exec(dateText)
      // get the 6 number parts as integers (lose fractions of second)
      let aParts = mCol.slice(1,6).map (s => parseInt(s,10))
      aParts[1] = aParts[1]-1 // Convert month Number to index
      return new Date(Date.UTC(...aParts))
    } else {
      return null
    }
  }
 
  
  const  getListOptionsForMethod = function (methodName ='GET', eTag = '*') {
    let options = {
      method: 'GET',
      headers: {
        Accept: 'application/json; odata=fullmetadata'
      }
    }
    switch (methodName) {
      case 'POST':
      case 'PUT':
        options.method = methodName
        options.headers['Content-Type'] = 'application/json'
        options.headers['Accept'] = 'application/json; odata=verbose'
        break
      case 'MERGE':
        options.method = 'POST'
        options.headers['Content-Type'] = 'application/json'
        options.headers['Accept'] = 'application/json; odata=verbose'
        options.headers['X-HTTP-Method'] = 'MERGE'
        options.headers['IF-MATCH'] = eTag
        break
      case 'DELETE':
        options.method = 'DELETE'
        options.headers['Content-Type'] = 'application/json; odata=verbose'
        options.headers['Accept'] = 'application/json; odata=verbose'
        options.headers['X-HTTP-Method'] = 'MERGE'
        options.headers['IF-MATCH'] = eTag
        break
      default:
        break
    }
    return options 
  }

  /**
   * @description query the site for meta information and current security tokens etc.
   * @param {string} siteBase - the URL to the SharePoint site when it is NOT the standard as configured in this library (default to standard if ommitted)
   * @returns {object}
 */
  const getContextInfoForSite = async function (siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
     // Get a new digest ...
    let sJsonVerbose = 'application/json; odata=verbose'
    let url = `${siteBase}/_api/contextinfo`
    let options = {
      method: 'POST',
      body: '{}',
      headers: {
        'Content-Type': sJsonVerbose,
        Accept: sJsonVerbose,
        'Content-Length': 2
      }
    }
    let resp = await fetch(url, options)
    let data = await resp.json()
    
    return data?.d.GetContextWebInformation || {}
     
  }

  /**
  * @description query the site for current security tokens to use in POST/PATCH requests etc.
  * @param {string} siteBase - the URL to the SharePoint site when it is NOT the standard as configured in this library (default to standard if ommitted)
  * @returns {string}
  */
  const getRequestDigest = async function (siteBase = SITE_BASE) {
    let webInfo = await getContextInfoForSite(siteBase) // NB: Did not use auto prefix site base here because it will get done in teh called func
    let newDigest = webInfo.FormDigestValue
    return newDigest 
  }

  /**
   * 
   * @param {string} listNameOrGuid - When the passed string is not structured like a SharePoint GUID it is assumed that the text is a list name
   * @param {integer} itemId 
   * @returns Promise<string> - OK When Delete success or FAIL in event of error
   */
  async function deleteItemInListWithId(listNameOrGuid, itemId, eTag = '*', siteBase = SITE_BASE){
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    let listNameCopy
    let urlListItemsBase
    try {
      if (reIsGuidText.test(listNameOrGuid)) {
        listNameCopy = 'having unique identifier: ' + listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      }  else {
        listNameCopy =  listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }
      let deleteUrl = `${urlListItemsBase}/items(${itemId})`
      let deleteOptions = getListOptionsForMethod('DELETE')
      deleteOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase)
      deleteOptions.headers['IF-MATCH'] = eTag
      let response = await fetch(deleteUrl, deleteOptions)
      if (response.status <= 204) {
        return  OK
      } else {
        throw new Error(`Delete response code indicates the delete failed in attempt to delete item #${itemId} from list ${listNameCopy}.\nCode: ${response.status}\n${response.statusText} `)
      }


    } catch (error){
      console.error(error)
      return FAIL
    }
  }

  /**
   * When reading FROM SharePoint
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {object[]} aList - an array of items from a SP REST response
   * @returns {object[]} - An array of logical objects for consumption by application layer
   */
  function convertSharePointListToLogicalFromFieldMap(mapFields, aList){
    const aEffectiveMap = MAP_P2L_COMMON.concat(mapFields) //prefix standard selections
   // const sharePointFieldsToKeep = R.map(mapPickFirst)(aEffectiveMap)
    const logicalFieldsToSet = R.map(mapPickSecond)(aEffectiveMap)

    const fnGetValueFromPhysicalMappedItem = ( index, item) =>{
      
      const propertyPath = aEffectiveMap[index][0] // From the SharePoint Internal "Source" column
      const slashPos = propertyPath.indexOf('/')
      
      if ( slashPos > -1) {
        return R.pathOr(null, propertyPath.split('/'), item)
      } else {
        return item[propertyPath]
      }
    }
    
    const fnMapper = itm=>{
      let out = {}
      logicalFieldsToSet.forEach((target,idx)=>out[target] = fnGetValueFromPhysicalMappedItem(idx,itm))
      out.Editor = Object.assign({}, itm.Editor)
      out.editedOn = new Date(R.pathOr(null,['Modified'], itm))
      out.editorName = R.pathOr('',['Editor','Title'], itm)
      return out
    }
    // Convert either an Array or a single object
    const aConverted =  aList instanceof Array ? aList.map(fnMapper) : fnMapper(aList)
    return aConverted

  }

  /**
   * When Writing TO SharePoint
   * INTERNAL maps logical names to source list SharePoint field names
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {*} aListOrItem either an array of objects or single object having logical property names that should be converted to a form suitable to update a SharePoint list record
   * @returns {object{} | object} - Array (or single Object)
   */
  function convertLogicalToSharePointListFromFieldMap(mapFields, aListOrItem){
    
    const aLookups = fnGetInternalNamesByTypeFromMap('Lookup', mapFields)
    const aLookupsMulti = fnGetInternalNamesByTypeFromMap('LookupMulti', mapFields)
    const aChoiceMulti = fnGetInternalNamesByTypeFromMap('ChoiceMulti', mapFields)
    const aUsers = fnGetInternalNamesByTypeFromMap('User', mapFields)
    const aUserMulti = fnGetInternalNamesByTypeFromMap('UserMulti', mapFields)
    const aExpandedSingles = aLookups.concat(aUsers)
    const aExpandedMulti = aLookupsMulti.concat(aUserMulti)
    const aLookupOrUserHandling = aLookups.concat(aLookupsMulti, aUsers, aUserMulti)


    const sharePointFieldsToSet = R.map(mapPickFirst)(mapFields)
    const logicalFieldsInput = mapFields.map(mapPickSecond)
    const fnMapItemToSharePoint = (itm)=>{
      let out = {}
      logicalFieldsInput.forEach((logicalField,idx) => {
        let sharePointFieldName = sharePointFieldsToSet[idx] 
        if (sharePointFieldName.indexOf('/')>-1) return; // Early exit for a looked up field definition
        if (itm[logicalField] instanceof Date) {
          out[sharePointFieldName] =  localDateToSharePoint(itm[logicalField])
        } else if ( aLookupOrUserHandling.includes(sharePointFieldName)) {
          if (aExpandedMulti.includes(sharePointFieldName)) {
            // NULL values need to be converted to empty array for multi value fields
            out[sharePointFieldName + 'Id'] = R.pluck('Id', R.pathOr([],[logicalField],itm))
          } else { // Single field reference
            console.assert(aExpandedSingles.includes(sharePointFieldName), `ASSERT TRAP: Expected the SharePoint name ${sharePointFieldName} to be in the list either User or Lookup types, but it wasn't!`)
            out[sharePointFieldName + 'Id'] = R.pathOr(null, [logicalField, 'Id'], itm) // picks up the Id value from {... propertyName: {Id: 123, Title: 'some Title'}}
          }
        } else if (aChoiceMulti.includes(sharePointFieldName)) {
          // Multi choice fields should never send a NULL value. replace any null lists with an empty array
          out[sharePointFieldName] = itm[logicalField] instanceof Array? itm[logicalField] : []
        }else {
          out[sharePointFieldName] = itm[logicalField]
        }
        
      })
      return out
    }
    if (aListOrItem instanceof Array) {
      return aListOrItem.map(fnMapItemToSharePoint)
    } else {
      return fnMapItemToSharePoint(aListOrItem)
    }
    
  }

  /**
   * @description Build a new object with all the expected properties in the target list with sensible defaults passed
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {object} defaultObjects - simple object with initialisation values for the logical output
   * @returns 
   */
  function createTemplateRecordFromMapAndPassedDefaultKeySet (mapFields, defaultObjects) {
    const logicalFieldsInput = mapFields.map(mapPickSecond)
    let oTemplate = {}
    logicalFieldsInput.forEach((prop,idx) => oTemplate = R.assoc(prop, null, oTemplate))
    return Object.assign(oTemplate, defaultObjects ||{})
  }
  /**
   * 
   * @param {string} listNameOrGuid - Guid or List Name
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {object} item - Simple item to be mapped to SharePoint structure and persisted in a lits
   * @returns 
   */
  async function updateOrCreateListItemUsingPhysicalToLogicalMap(listNameOrGuid, mapFields, item, siteBase = SITE_BASE, useExtendedReturn = false) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    let listNameCopy
    let isUpdateOperation = false // Assume a create at this stage
    let urlListItemsBase
    try {
      if (typeof item === 'object') {
        isUpdateOperation = item.hasOwnProperty('Id')
      } else {
        throw new Error(`An item object is required!`)
      }
      
      if (reIsGuidText.test(listNameOrGuid)) {
        listNameCopy = 'having unique identifier: ' + listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      }  else {
        listNameCopy =  listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }
      let postUrl = urlListItemsBase + '/items'
      if (isUpdateOperation) postUrl += `(${item.Id})`;

      let itemListSafe = convertLogicalToSharePointListFromFieldMap(mapFields, item) // becomes file_1, field_2 etc
      let eTag = item?.eTag || '*'
      let postBody = JSON.stringify(itemListSafe)
      let postOptions = isUpdateOperation ?  getListOptionsForMethod('MERGE', eTag) :  getListOptionsForMethod('POST')
      postOptions.body = postBody
      postOptions.headers['Content-Length'] = postBody.length
      postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase)

      if (useExtendedReturn) {
        postOptions.headers.Accept = 'application/json; odata=fullmetadata'
      }

      let response = await fetch(postUrl, postOptions)
      if (useExtendedReturn) {
        if (isUpdateOperation){
          // Don't know why, but an update seems not to return the physical JSON object
          // console.warn('re-read updated item via',postUrl)
          let respReadUpdated = await fetch(postUrl, getListOptionsForMethod('GET'))
          try {
            return await respReadUpdated.json()
          } catch(ex) {
            console.error('Failed to read the updated item from the API\nError:' + error.message)
            return item //
          }
          
        } else {
          return await response.json()
        }
      }
      else if (response.status <= 204) {
        return OK
      } else {
        console.error(response)
        cacheFailedRequests.push(response)
        return `${FAIL}\n${response.statusText}`
      }
      
    } catch (error) {
      console.error(`Failed to get items from list "${listNameCopy}"\n${error.message}`)
      return []      
    }
  }

  /**
   * 
   * @param {string} listNameOrGuid - Guid or List Name
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {object} item - Simple item to be mapped to SharePoint structure and persisted in a lits
   * @returns 
   */
  async function createListItemUsingPhysicalToLogicalMapFullResponse(listNameOrGuid, mapFields, item, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    let listNameCopy
    
    let urlListItemsBase
    try {
      if (typeof item !== 'object') {
        throw new Error(`An item object is required!`)
      }
      
      if (reIsGuidText.test(listNameOrGuid)) {
        listNameCopy = 'having unique identifier: ' + listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      }  else {
        listNameCopy =  listNameOrGuid
        urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }
      let postUrl = urlListItemsBase + '/items'
      
      let itemListSafe = convertLogicalToSharePointListFromFieldMap(mapFields, item) // becomes file_1, field_2 etc
      let postBody = JSON.stringify(itemListSafe)
      let postOptions =  getListOptionsForMethod('POST')
      postOptions.headers.Accept = 'application/json; odata=fullmetadata' // Override the usual "nometadata"
      postOptions.body = postBody
      postOptions.headers['Content-Length'] = postBody.length
      postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase)

      let response = await fetch(postUrl, postOptions)
      if (response.status <= 204) {
        return await response.json()
      } else {
        console.error(response)
        cacheFailedRequests.push(response)
        return `${FAIL}\n${response.statusText}`
      }
      
    } catch (error) {
      console.error(`Failed to get items from list "${listNameCopy}"\n${error.message}`)
      return []      
    }
  }
  
  /**
   * Internal
   * Construct the REST query parameters for some given passed options
   * @param {string} options.filterClause - Restrictions on retrieval based on PHYSICAL (SharePoint)  field names
   * @param {string} options.selectClause - List of fields to return from the SharePoint list. Must use PHYSICAL (SharePoint)  field names
   * @param {string} options.orderByClause - List of fields to sort the result by
   * @param {number} options.itemLimit - Max Number of items to retrieve
   * @param {string} options.oDataNextLink - when using a "paged" request from an original (or subsequent) query with items remaining
   
   * @returns 
   */
  const getSiteListQueryFromOptions = function (options={itemLimit:5000}, mapFields){
    let {itemLimit, filterClause, selectClause, orderByClause,oDataNextLink} = options

    // Get from page result - repeated the query used previously
    if (oDataNextLink) {
      return oDataNextLink.substring(oDataNextLink.indexOf('?') )
    }
    
    let urlQuery = `?$top=${itemLimit ? itemLimit : 5000}`
    if (filterClause) {
      urlQuery += '&$filter=' + filterClause
    }
    if (selectClause) {
      urlQuery += '&$select=' + selectClause
      if (selectClause.indexOf('Editor/Title') === -1) urlQuery += ',Editor/Title'; // Whether asked for or not we will return Editor Name
      if (selectClause.indexOf('Editor/EMail') === -1) urlQuery += ',Editor/EMail'; //  ditto for email
      const expandParam = '&$expand='
      let expandClause = R.compose( // READ BOTTOM UP ^^
        R.reduce((acc, cur)=> {return acc + (acc.length > expandParam.length ? ',':'') + cur}, expandParam), // assemble URL param
        R.uniq, // remove duplicates
        R.append('Editor'), // Always need editor!
        R.map(s => s.substring(0, s.indexOf('/'))), // pick left side of first slash (should never be more than 1)
        R.filter(s=> s.indexOf('/')>-1), // keep those with a path slash
        R.map(R.trim), // trim space around field names
        R.split(',') // convert clause to array of parts
      )(selectClause)

      urlQuery += expandClause // At very least this will be Editor!

    } else {
      // any lookups, User or UserMulti fields?
      let genericExpandClause = '&$expand=Editor'
      let genericSelect = '&$select=*,Editor/Title,Editor/EMail'
      let aLookupsSingle = fnGetInternalNamesByTypeFromMap ('Lookup', mapFields)
      let aLookupsMulti = fnGetInternalNamesByTypeFromMap ('LookupMulti', mapFields)
      let aLookups = aLookupsSingle.concat(aLookupsMulti)
      if (aLookups.length) { // Lookups need lots of extra processing to form a query
        
        // process any lookup internal names that contains a slash to get just the 1st path part
        let aLookupExpands = R.compose(
          R.uniq,
          R.map(fld => fld.indexOf('/') > - 1 ? fld.substring(0, fld.indexOf('/')): fld)
        )(aLookups)
        genericExpandClause += ','+ aLookupExpands.join(',')
        
        let aLookupSelects = R.map(fld => fld.indexOf('/') > - 1 ? fld : fld.substring(0, fld.indexOf('/')) +'Title')(aLookups)
        
        // As a minimum, we need <expandName>/Id and <expandName>/Title
        let aLookupMinimalSelects = R.map(key => `${key}/Id,${key}/Title`)(aLookupExpands)
        aLookupSelects = aLookupSelects.concat(aLookupMinimalSelects)
        
        genericSelect += ',' + R.compose(R.join(','), R.uniq)(aLookupSelects)
      }
      let aUser = fnGetInternalNamesByTypeFromMap ('User', mapFields)
      let aUserMulti = fnGetInternalNamesByTypeFromMap ('UserMulti', mapFields)
      let aAllUserFields = aUser.concat(aUser, aUserMulti)
      if (aAllUserFields.length){
        genericExpandClause += ',' + aAllUserFields.join(',')
        genericSelect +=  ',' + aAllUserFields.map(colName => `${colName}/Id,${colName}/Title,${colName}/EMail`).join(',')       
      }
      
      

      urlQuery += genericSelect + genericExpandClause // When not asked for specific give all (*) plus the Editor/Title+EMail
    }
    if (orderByClause) {
      urlQuery += '&$orderby=' + orderByClause
    }
    return urlQuery
  }
  

  /**
   * 
   * @param {string} listNameOrGuid - Either List Name or its internal GUID (without curly braces!)
   * @param {string[][]} mapFields - 2D array where 1st entry is the internal SharePoint list name of a field, 2nd is the logical application name and (optional) 3rd is a display name
   * @param {object} options - Example {selectClause: 'Title,someValue,Field_1', orderBy: 'Modified desc', filterBy: 'startswith(Title,\'A\'')}
   * @param {string} options.filterClause - Restrictions on retrieval based on PHYSICAL (SharePoint)  field names
   * @param {string} options.selectClause - List of fields to return from the SharePoint list. Must use PHYSICAL (SharePoint)  field names
   * @param {string} options.orderByClause - List of fields to sort the result by
   * @param {number} options.itemLimit - Max Number of items to retrieve
   * @returns Promise<object[]>
   */
  const getSiteListItems = async function (listNameOrGuid, mapFields, options={itemLimit:5000}, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    let resp
    let data
    let url
    try {
      let urlListBase = ''
      if (reIsGuidText.test(listNameOrGuid)) {
        urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      } else {
        urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }
      url = `${urlListBase}/items${getSiteListQueryFromOptions(options, mapFields)}`
      resp = await fetch (url, getListOptionsForMethod('GET'))
      data = await resp.json()
      let aOut 
      if (options.raw) {
        aOut = fGetItemArrayFromRestData(data)
        if (data.hasOwnProperty('odata.nextLink')) aOut.oDataNextLink = data['odata.nextLink']
      
      } else {
        aOut =  convertSharePointListToLogicalFromFieldMap(mapFields,  fGetItemArrayFromRestData(data))
        if (data.hasOwnProperty('odata.nextLink')) aOut.oDataNextLink = data['odata.nextLink'] 
      } 
      return aOut   // Note; Bad practice to modify "standard" object properties, but needed to shove the data['odata.nextLink'] somewhere to be usefult
    } catch (error) {
      console.error(`Failure getting site list items.\nError: ${error instanceof Error ? error.message : error?.toString()}`)
      console.log(`url = ${url}`)
      console.log(`response status = ${ resp ? resp.status : '-'}`)
      console.log(`response status text= ${ resp ? resp.statusText : '-'}`)
      if (resp) {

        cacheFailedRequests.push(resp)
      }
    }
  }
  /**
   * 
   * @param {*} listNameOrGuid 
   * @param {number} itemId - record number
   * @param {File} file - typically from a File input element
   * @param {*} siteBase - url to override the default site base (optional)
   * @returns JSON formatted metadata about teh uploaded file (or an error!)
   */
  const addAttachmentToListItem = async function (listNameOrGuid, itemId, file, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    try {
      let urlListBase = ''
      if (reIsGuidText.test(listNameOrGuid)) {
        urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      } else {
        urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }
      let url = `${urlListBase}/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`
      let postOptions = getListOptionsForMethod('POST')
      postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase)
      delete postOptions.headers['Content-Type'] // let the browser set this from the file
      postOptions = Object.assign({processData: false, body:file },postOptions)
      let result = await fetch(url,postOptions)
      return await result.json()
    } catch (error) {
      console.error(error)
      return error
    }
  }
  const getAttachmentFromListItemNameOrIndex = async function (listNameOrGuid, itemId, fileNameOrIndex, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    try {
      let urlListBase = ''
      if (reIsGuidText.test(listNameOrGuid)) {
        urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      } else {
        urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }

      let url = `${urlListBase}/items(${itemId})/AttachmentFiles`
      let options = getListOptionsForMethod('GET')
      delete options.headers['Content-Type'] 
      if (typeof fileNameOrIndex === 'number'){
        if (Math.floor(fileNameOrIndex) !== fileNameOrIndex || fileNameOrIndex < 0) {
          throw new Error('fileNameOrIndex must be an integer zero or greater')
        }
        let resp = await fetch(url, options)
        let data = await resp.json()
        if (data.value.length === 0) {
          throw new Error(`No attachments found for item ${itemId} in list ${listNameOrGuid}`)
        } else if (fileNameOrIndex >= data.value.length) {
          throw new Error(`Requested attachment index ${fileNameOrIndex} is out of range for item ${itemId} in list ${listNameOrGuid}`)
        } else {
          url += `('${encodeURIComponent(data.value[fileNameOrIndex].FileName)}')/$value`
          return  fetch(url, options)
        }
      } else if (typeof fileNameOrIndex === 'string'){
        url += `('${encodeURIComponent(fileNameOrIndex)}')/$value`
        return  fetch(url, options)

      } else {
        throw new Error('fileNameOrIndex must be a number or string')
      }
      
    } catch (error) {
      console.error(error)
      return error
    }
  }
  const deleteAttachmentFromListItemNameOrIndex = async function (listNameOrGuid, itemId, fileNameOrIndex, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    try {
      let urlListBase = ''
      if (reIsGuidText.test(listNameOrGuid)) {
        urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase)
      } else {
        urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase)
      }

      let url = `${urlListBase}/items(${itemId})/AttachmentFiles`
      let options = getListOptionsForMethod('DELETE')
      delete options.headers['Content-Type'] 
      options.headers['X-RequestDigest'] = await getRequestDigest(siteBase)
      options.headers['X-HTTP-Method'] = 'DELETE'
      if (typeof fileNameOrIndex === 'number'){
        if (Math.floor(fileNameOrIndex) !== fileNameOrIndex || fileNameOrIndex < 0) {
          throw new Error('fileNameOrIndex must be an integer zero or greater')
        }
        let resp = await fetch(url, options)
        let data = await resp.json()
        if (data.value.length === 0) {
          throw new Error(`No attachments found for item ${itemId} in list ${listNameOrGuid}`)
        } else if (fileNameOrIndex >= data.value.length) {
          throw new Error(`Requested attachment index ${fileNameOrIndex} is out of range for item ${itemId} in list ${listNameOrGuid}`)
        } else {
          url += `('${encodeURIComponent(data.value[fileNameOrIndex].FileName)}')`
          return  fetch(url, options)
        }
      } else if (typeof fileNameOrIndex === 'string'){
        url += `('${encodeURIComponent(fileNameOrIndex)}')`
        return  fetch(url, options)

      } else {
        throw new Error('fileNameOrIndex must be a number or string')
      }
    
    } catch (error) {
      console.error(error)
      return error
    }
  }


  /**
   * @description Used by application layer to construct option objects for filterClause etc
   * Recommend to use partial application (function bind) to make an application layer specif function like fn(string)=>string
   * @param {*} aMapTuple 
   * @param {string} logicalName - the name of the field used in code 
   * @returns {string} the name used in SharePoint for the passed logical name or undefined when the name is not found
   */
  function getLogicalToPhysicalFieldNameForMap (aMapTuple, logicalName) {
    let aPhysicalName = aMapTuple.map(mapPickFirst)
    let aLogicalName = aMapTuple.map(mapPickSecond)
    let idx = aLogicalName.findIndex(R.equals(logicalName))
    if (idx > -1) {
      return aPhysicalName[idx]
    }
  }

  class ObjectVersionDifference {
    logicalProperty;
    caption;
    editedValue;
    newValue;
    retain;
    isSystem;
    constructor(logicalProperty, caption, editedValue, newValue) {
      let reIsSystemField = /\(system\)/ig
      this.logicalProperty = logicalProperty
      this.caption = caption
      this.editedValue = editedValue
      this.newValue = newValue
      this.retain = 'OLD'
      this.isSystem = reIsSystemField.test(caption)
    }
    
  }
  function compareMappedEditedToLatestLogicalObjects (mapToUse, edited, latest) {
    let aDifferences = [] // Array of difference objects

    let aLogicalName = mapToUse.map( mapPickSecond)
    let aCaptions = mapToUse.map(mapPickThird)
    
    aLogicalName.forEach((logical, idx) =>{
      let editedValue = edited[logical]
      let latestValue = latest[logical]
      let hasDifferences = false
      if (typeof editedValue !== typeof latestValue) {
        hasDifferences = true
      } else if (R.not(R.equals(editedValue,latestValue))) {
        hasDifferences = true        
      }
      if (hasDifferences) {
        let differenceSpecification = new ObjectVersionDifference(logical, aCaptions[idx], editedValue, latestValue)
        if (differenceSpecification.isSystem) {
          differenceSpecification.retain = 'NEW' // assume that the EDITED field should not (by default) overwrite a SYSTEM field that is found
        }
        aDifferences.push(differenceSpecification)
      }
    })

    return aDifferences

  }

  const getCurrentUserInSite =  async function (siteBase = SITE_BASE) {
    if (!cacheCurrentUserInSite[siteBase]) {
      let resp = await fetch(`${siteBase}/_api/web/currentUser`, getListOptionsForMethod('GET'))
      let data = await resp.json()
      cacheCurrentUserInSite[siteBase] = data
    }
    return cacheCurrentUserInSite[siteBase]
  }

  const getGroupAndMembersByNameOrIdInSite  = async function (nameOrId, siteBase = SITE_BASE) {
    siteBase = autoPrefixSiteUrlWithBase(siteBase) 
    const urlGroup = `${siteBase}/_api/web/SiteGroups/GetBy${typeof nameOrId === 'number' ? `ID(${nameOrId})` : `Name('${encodeURIComponent(nameOrId)}')`}`
    const urlMembers = urlGroup + '/users'
    const options = getListOptionsForMethod('GET')
    try {
      const [oGroup, oUserQuery] = await Promise.all([
        fetch (urlGroup, options).then (r => r.json()),
        fetch (urlMembers, options).then (r => r.json()),
      ])
      return R.assoc('members', R.pathOr([],['value'], oUserQuery), oGroup)
      
    } catch (ex) {
      return {message: ex instanceof Error ? ex.message : ex?.toString()}
    }

  }

  function getSemVer() {let parts = SemVer.split('.'); return {major: parts[0], minor: parts[1], revision: parts[2]}}
  
/**
 * Expose some of the "internal" utility functions for future flexibility and testing
 */
  
  
  namespaceObject.getSiteListItems = getSiteListItems // [Read]
  namespaceObject.updateOrCreateListItemUsingPhysicalToLogicalMap = updateOrCreateListItemUsingPhysicalToLogicalMap // [Create] and [Update]
  namespaceObject.createListItemUsingPhysicalToLogicalMapFullResponse = createListItemUsingPhysicalToLogicalMapFullResponse // [Create] only
  namespaceObject.deleteItemInListWithId = deleteItemInListWithId // [Delete]
  
  namespaceObject.getLogicalToPhysicalFieldNameForMap = getLogicalToPhysicalFieldNameForMap

  namespaceObject.getRequestDigest = getRequestDigest
  namespaceObject.createTemplateRecordFromMapAndPassedDefaultKeySet = createTemplateRecordFromMapAndPassedDefaultKeySet

  namespaceObject.convertIsoDateTextToDate = fConvertIsoDateTextToDate
  namespaceObject.localDateToSharePoint = localDateToSharePoint

  
  namespaceObject.compareMappedEditedToLatestLogicalObjects = compareMappedEditedToLatestLogicalObjects
  
  namespaceObject.getCurrentUserInSite = getCurrentUserInSite
  namespaceObject.getGroupAndMembersByNameOrIdInSite = getGroupAndMembersByNameOrIdInSite 

  namespaceObject.popLastFailedRequest =  cacheFailedRequests.pop
  namespaceObject.getCountFailedResponse = () => cacheFailedRequests.length

  namespaceObject.convertLogicalToSharePointListFromFieldMap = convertLogicalToSharePointListFromFieldMap
  namespaceObject.convertSharePointListToLogicalFromFieldMap = convertSharePointListToLogicalFromFieldMap
  namespaceObject.getContextInfoForSite = getContextInfoForSite

  namespaceObject.getSemVer = getSemVer
  namespaceObject.getSiteBase = getSiteBase
  namespaceObject.setSiteBase = setSiteBase

  namespaceObject.addAttachmentToListItem = addAttachmentToListItem
  namespaceObject.getAttachmentFromListItemNameOrIndex = getAttachmentFromListItemNameOrIndex
  namespaceObject.deleteAttachmentFromListItemNameOrIndex = deleteAttachmentFromListItemNameOrIndex
  return namespaceObject
})(
  globalThis.maceSPListUtility || {} // ~~CONFIGURE HERE~~ Set the globalName required
)
