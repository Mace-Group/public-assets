/**
 * This script creates an old fashioned "module" (or namespace) pattern object
 * It provides utility functions to help determine users and groups in a SharePoint site
 * 
 * Designers note
 * The code here is "portable" because it contains no names and/or IDs of any site
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
 */

var SITE_BASE=''

// ~~CONFIGURE HERE~~ set the global name required
globalThis.maceSPUserUtility = (function (namespaceObject) {
  'use strict'
  console.log('Loading Module maceSPUserUtility.js') // This is useful to trace any issues with load sequence
  const SemVer='2.1.2'

  /**
   * The SharePoint lists managed by this module were mainly created by the SharePoint List from Spreadsheet facility
   * The result is that many of the internal field names are unsuited to logical coding conventions (i.e field_1, field_2 etc)
   * Consequently this library has a translation capability used to map between the obscure internals return from the REST API to
   * useful and consistent names for the consumers of THIS library.
   */
  
  
  
  if (document?.currentScript) {
    let src = document.currentScript.src
    let reFindSite = /(.+)\/siteAssets\/(.*)/i
    if (reFindSite.test(src)){
      let match = reFindSite.exec(src)
      console.log(`maceSPUserUtility dynamically determined its site base to be ${match[1]}`)
      SITE_BASE = match[1]
    }
  }
  let cacheCurrentUserInSite = {}
  const OK = 'OK'
  const FAIL = 'FAIL'
    
  // Upscale Excel to List Mapping conversions
  // The following maps convert SharePoint List internal IDs to object properties to make coding life easier!
  
  

  // Help problem diagnosis with a list of dependencies that THIS module will use
  const assumedLoadedGlobals = [
      {globalName: 'R', description: 'Ramda 0.29+'},
  
  ]
  assumedLoadedGlobals.forEach(m => {
      if (!globalThis[m.globalName]) {
      console.warn(`Assumed global "${m.globalName}" for ${m.description} is not loaded. Errors will be likely.`)
      }
  })
  
  
  
  // Utility functions 
  
  const fGetItemArrayFromRestData = R.pathOr([],['value']) // An implicitly curried fn call just needs a REST/JSON response
  
  const getGroupApiRelativeUrlFromNameOrId = (nameOrId) => `_api/web/SiteGroups/GetBy${typeof nameOrId === 'number' ? `ID(${nameOrId})` : `Name('${encodeURIComponent(nameOrId)}')`}`
  
    
  function autoPrefixSiteUrlWithBase (siteBase) {
      if (typeof siteBase === 'string' && !siteBase.startsWith('https://')) {
      return `${SITE_BASE}${ siteBase.startsWith('/') ? '' : '/'}${siteBase}` //prefix the argument with <standard Site Url>/
      } else {
      return siteBase // They passed something weird so give it back!
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
    const urlGroup = `${siteBase}/${getGroupApiRelativeUrlFromNameOrId(nameOrId)}`
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
  
  /**
   * Enumerates all the site groups that have Enumerable members (i.e. filters out all the groups used from SharingLinks etc.)
   * @param {string} siteBase  - URL of site to examine
   */
  const getAllSiteGroups = async function (siteBase = SITE_BASE) {
    // let url = `${siteBase}/_api/web/SiteGroups?$select=Id,Title,Description,OwnerTitle&$filter=OnlyAllowMembersViewMembership eq false`
    let url = `${siteBase}/_api/web/SiteGroups?$select=Id,Title,Description,OwnerTitle`
    let aGroups = await fetch(url, getListOptionsForMethod('GET')).then(r=> r.json()).then(d=> fGetItemArrayFromRestData(d))
    return aGroups
  }

  const getGroupMembersByGroupNameOrId = async function (nameOrId, siteBase = SITE_BASE) {
    let url = `${siteBase}/${getGroupApiRelativeUrlFromNameOrId(nameOrId)}/users`  
    let aMembers = fGetItemArrayFromRestData(await fetch(url, getListOptionsForMethod('GET')).then(r=>r.json()))
    return aMembers
  }
  /**
   * Enumerates all the site uses by their group membership
   * There will be MANY fails for internal groups created by sharing documents etc.
   * @param {string} usersGroupsOrBoth - Filter the result to either [U] Users, [G] Groups or [B] Both
   * @param {string} siteBase  - URL of site to examine
   */
  const getAllSiteMembers = async function (usersGroupsOrBoth = 'U',siteBase = SITE_BASE) {
    let aGroups = await getAllSiteGroups(siteBase)
    let aGroupMemberSets = await Promise.all(aGroups.map(grp => getGroupMembersByGroupNameOrId(grp.Id, siteBase)))
    let fnFilter
    switch (usersGroupsOrBoth.substring(0,1).toUpperCase()) {
      case 'U':
        fnFilter = (obj) => obj?.UserId !== null
        break;
      case 'G':
        fnFilter = (obj) => obj?.UserId === null
        break;
      default: 
        fnFilter = R.T // The TRUE function (to not filter anything)
    }
    let aMapPeople= R.compose(  // It's a "compose" so read RTL (i.e. bottom up!)
      R.sortBy(R.view(R.lensProp('Title'))),
      R.uniqBy(R.view(R.lensProp('Id'))),
      R.filter(fnFilter),
      R.flatten
    ) (aGroupMemberSets)
    

    return aMapPeople
  }

  const getAllSiteUsers = async function (filters={email:'', title:''}, siteBase = SITE_BASE) {
   let url = `${siteBase}/_api/web/SiteUsers?$select=Id,Title,Email` 
   let resp = await fetch(url, getListOptionsForMethod('GET'))
   let data = await resp.json()
   let aUsers = R.pathOr([],['value'], data)
   if (filters?.email) {
    let reEmail = new RegExp(filters.email,"i")
    aUsers = aUsers.filter(usr => reEmail.test(usr.Email))
   }
   if (filters?.title) {
    let reName = new RegExp(filters.title,"i")
    aUsers = aUsers.filter(usr => reName.test(usr.Title))
   }
   return aUsers
  }

/**
 * Expose some of the "internal" utility functions for future flexibility and testing
 */
    
  
    
    namespaceObject.getCurrentUserInSite = getCurrentUserInSite
    namespaceObject.getGroupAndMembersByNameOrIdInSite = getGroupAndMembersByNameOrIdInSite 
    namespaceObject.getGroupMembersByGroupNameOrId = getGroupMembersByGroupNameOrId
    namespaceObject.getAllSiteGroups = getAllSiteGroups
    namespaceObject.getAllSiteMembers = getAllSiteMembers
    namespaceObject.getAllSiteUsers = getAllSiteUsers

    namespaceObject.getBaseUrl = () => SITE_BASE
    namespaceObject.setBaseUrl = (url) => SITE_BASE = url
  
    return namespaceObject
  })(
  globalThis.maceSPUserUtility || {} // ~~CONFIGURE HERE~~ Set the globalName required
)