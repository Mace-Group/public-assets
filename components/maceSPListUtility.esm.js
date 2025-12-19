// ESM version of maceSPListUtility
import * as R from 'https://cdn.jsdelivr.net/npm/ramda@0.32.0/+esm'
import moment from 'https://cdn.jsdelivr.net/npm/moment@2.30.1/+esm'

const SemVer = '2.1.1';

let SITE_BASE;
if (document?.currentScript) {
  let src = document.currentScript.src;
  let reFindSite = /(.+)\/siteAssets\/(.*)/i;
  if (reFindSite.test(src)){
    let match = reFindSite.exec(src);
    console.log(`maceSPListUtility dynamically determined its site base to be ${match[1]}`);
    SITE_BASE = match[1];
  }
}

let cacheCurrentUserInSite = {};
const OK = 'OK';
const FAIL = 'FAIL';
const cacheFailedRequests = [];

function getSiteBase() { return SITE_BASE; }
function setSiteBase(url) { SITE_BASE = url; }

const MAP_P2L_COMMON = [
  ['Modified', 'Modified', 'Modified Date'],
  ['Id','Id', 'Internal ID', 'SharePoint List Id'],
  ['odata.etag', 'eTag', 'Edit Tag'],
  ['odata.id', 'oDataId', 'OData Identifier']
];

const mapPickFirst = R.view(R.lensIndex(0));
const mapPickSecond = R.view(R.lensIndex(1));
const mapPickThird = R.view(R.lensIndex(2));
const mapPickFourth = R.view(R.lensIndex(3));
const fnGetInternalNamesByTypeFromMap = (typeName, mapArray) => mapArray.reduce((acc,cur)=> {if(cur[3] === typeName) acc.push(cur[0]); return acc}, []);
const fGetItemArrayFromRestData = R.pathOr([],['value']);
const reIsISODateText = /(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(\.\d{1,})?)([Z+-]\d{0,4})/i;
const reIsGuidText =/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const localDateToSharePoint = function (value) { return moment(value).toISOString().substring(0,19) + 'Z'; };

function autoPrefixSiteUrlWithBase(siteBase) {
  if (typeof siteBase === 'string' && !siteBase.startsWith('https://')) {
    return `${SITE_BASE}${ siteBase.startsWith('/') ? '' : '/'}${siteBase}`;
  } else {
    return siteBase;
  }
}

const getSiteListRestUrlByName = function (listName, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  return `${siteBase}/_api/web/lists/GetByTitle('${listName}')`;
};

const getSiteListRestUrlByGuid = function (guid, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  return `${siteBase}/_api/web/lists(guid'${guid}')`;
};

const fConvertIsoDateTextToDate = function (dateText) {
  if (reIsISODateText.test(dateText)){
    let mCol = reIsISODateText.exec(dateText);
    let aParts = mCol.slice(1,6).map (s => parseInt(s,10));
    aParts[1] = aParts[1]-1;
    return new Date(Date.UTC(...aParts));
  } else {
    return null;
  }
};

const getListOptionsForMethod = function (methodName ='GET', eTag = '*') {
  let options = {
    method: 'GET',
    headers: {
      Accept: 'application/json; odata=fullmetadata'
    }
  };
  switch (methodName) {
    case 'POST':
    case 'PUT':
      options.method = methodName;
      options.headers['Content-Type'] = 'application/json';
      options.headers['Accept'] = 'application/json; odata=verbose';
      break;
    case 'MERGE':
      options.method = 'POST';
      options.headers['Content-Type'] = 'application/json';
      options.headers['Accept'] = 'application/json; odata=verbose';
      options.headers['X-HTTP-Method'] = 'MERGE';
      options.headers['IF-MATCH'] = eTag;
      break;
    case 'DELETE':
      options.method = 'DELETE';
      options.headers['Content-Type'] = 'application/json; odata=verbose';
      options.headers['Accept'] = 'application/json; odata=verbose';
      options.headers['X-HTTP-Method'] = 'MERGE';
      options.headers['IF-MATCH'] = eTag;
      break;
    default:
      break;
  }
  return options;
};

const getContextInfoForSite = async function (siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  let sJsonVerbose = 'application/json; odata=verbose';
  let url = `${siteBase}/_api/contextinfo`;
  let options = {
    method: 'POST',
    body: '{}',
    headers: {
      'Content-Type': sJsonVerbose,
      Accept: sJsonVerbose,
      'Content-Length': 2
    }
  };
  let resp = await fetch(url, options);
  let data = await resp.json();
  return data?.d.GetContextWebInformation || {};
};

const getRequestDigest = async function (siteBase = SITE_BASE) {
  let webInfo = await getContextInfoForSite(siteBase);
  let newDigest = webInfo.FormDigestValue;
  return newDigest;
};

async function deleteItemInListWithId(listNameOrGuid, itemId, eTag = '*', siteBase = SITE_BASE){
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  let listNameCopy;
  let urlListItemsBase;
  try {
    if (reIsGuidText.test(listNameOrGuid)) {
      listNameCopy = 'having unique identifier: ' + listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    }  else {
      listNameCopy =  listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let deleteUrl = `${urlListItemsBase}/items(${itemId})`;
    let deleteOptions = getListOptionsForMethod('DELETE');
    deleteOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase);
    deleteOptions.headers['IF-MATCH'] = eTag;
    let response = await fetch(deleteUrl, deleteOptions);
    if (response.status <= 204) {
      return  OK;
    } else {
      throw new Error(`Delete response code indicates the delete failed in attempt to delete item #${itemId} from list ${listNameCopy}.\nCode: ${response.status}\n${response.statusText} `);
    }
  } catch (error){
    console.error(error);
    return FAIL;
  }
}

function convertSharePointListToLogicalFromFieldMap(mapFields, aList){
  const aEffectiveMap = MAP_P2L_COMMON.concat(mapFields);
  const logicalFieldsToSet = R.map(mapPickSecond)(aEffectiveMap);

  const fnGetValueFromPhysicalMappedItem = ( index, item) =>{
    const propertyPath = aEffectiveMap[index][0];
    const slashPos = propertyPath.indexOf('/');
    if ( slashPos > -1) {
      return R.pathOr(null, propertyPath.split('/'), item);
    } else {
      return item[propertyPath];
    }
  };

  const fnMapper = itm=>{
    let out = {};
    logicalFieldsToSet.forEach((target,idx)=>out[target] = fnGetValueFromPhysicalMappedItem(idx,itm));
    out.Editor = Object.assign({}, itm.Editor);
    out.editedOn = new Date(R.pathOr(null,['Modified'], itm));
    out.editorName = R.pathOr('',['Editor','Title'], itm);
    return out;
  };
  const aConverted =  aList instanceof Array ? aList.map(fnMapper) : fnMapper(aList);
  return aConverted;
}

function convertLogicalToSharePointListFromFieldMap(mapFields, aListOrItem){
  const aLookups = fnGetInternalNamesByTypeFromMap('Lookup', mapFields);
  const aLookupsMulti = fnGetInternalNamesByTypeFromMap('LookupMulti', mapFields);
  const aChoiceMulti = fnGetInternalNamesByTypeFromMap('ChoiceMulti', mapFields);
  const aUsers = fnGetInternalNamesByTypeFromMap('User', mapFields);
  const aUserMulti = fnGetInternalNamesByTypeFromMap('UserMulti', mapFields);
  const aExpandedSingles = aLookups.concat(aUsers);
  const aExpandedMulti = aLookupsMulti.concat(aUserMulti);
  const aLookupOrUserHandling = aLookups.concat(aLookupsMulti, aUsers, aUserMulti);

  const sharePointFieldsToSet = R.map(mapPickFirst)(mapFields);
  const logicalFieldsInput = mapFields.map(mapPickSecond);
  const fnMapItemToSharePoint = (itm)=>{
    let out = {};
    logicalFieldsInput.forEach((logicalField,idx) => {
      let sharePointFieldName = sharePointFieldsToSet[idx];
      if (sharePointFieldName.indexOf('/')>-1) return;
      if (itm[logicalField] instanceof Date) {
        out[sharePointFieldName] =  localDateToSharePoint(itm[logicalField]);
      } else if ( aLookupOrUserHandling.includes(sharePointFieldName)) {
        if (aExpandedMulti.includes(sharePointFieldName)) {
          out[sharePointFieldName + 'Id'] = R.pluck('Id', R.pathOr([],[logicalField],itm));
        } else {
          out[sharePointFieldName + 'Id'] = R.pathOr(null, [logicalField, 'Id'], itm);
        }
      } else if (aChoiceMulti.includes(sharePointFieldName)) {
        out[sharePointFieldName] = itm[logicalField] instanceof Array? itm[logicalField] : [];
      }else {
        out[sharePointFieldName] = itm[logicalField];
      }
    });
    return out;
  };
  if (aListOrItem instanceof Array) {
    return aListOrItem.map(fnMapItemToSharePoint);
  } else {
    return fnMapItemToSharePoint(aListOrItem);
  }
}

function createTemplateRecordFromMapAndPassedDefaultKeySet (mapFields, defaultObjects) {
  const logicalFieldsInput = mapFields.map(mapPickSecond);
  let oTemplate = {};
  logicalFieldsInput.forEach((prop,idx) => oTemplate = R.assoc(prop, null, oTemplate));
  return Object.assign(oTemplate, defaultObjects ||{});
}

async function updateOrCreateListItemUsingPhysicalToLogicalMap(listNameOrGuid, mapFields, item, siteBase = SITE_BASE, useExtendedReturn = false) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  let listNameCopy;
  let isUpdateOperation = false;
  let urlListItemsBase;
  try {
    if (typeof item === 'object') {
      isUpdateOperation = item.hasOwnProperty('Id');
    } else {
      throw new Error(`An item object is required!`);
    }
    if (reIsGuidText.test(listNameOrGuid)) {
      listNameCopy = 'having unique identifier: ' + listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    }  else {
      listNameCopy =  listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let postUrl = urlListItemsBase + '/items';
    if (isUpdateOperation) postUrl += `(${item.Id})`;

    let itemListSafe = convertLogicalToSharePointListFromFieldMap(mapFields, item);
    let eTag = item?.eTag || '*';
    let postBody = JSON.stringify(itemListSafe);
    let postOptions = isUpdateOperation ?  getListOptionsForMethod('MERGE', eTag) :  getListOptionsForMethod('POST');
    postOptions.body = postBody;
    postOptions.headers['Content-Length'] = postBody.length;
    postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase);

    if (useExtendedReturn) {
      postOptions.headers.Accept = 'application/json; odata=fullmetadata';
    }

    let response = await fetch(postUrl, postOptions);
    if (useExtendedReturn) {
      if (isUpdateOperation){
        let respReadUpdated = await fetch(postUrl, getListOptionsForMethod('GET'));
        try {
          return await respReadUpdated.json();
        } catch(ex) {
          console.error('Failed to read the updated item from the API\nError:' + error.message);
          return item;
        }
      } else {
        return await response.json();
      }
    }
    else if (response.status <= 204) {
      return OK;
    } else {
      console.error(response);
      cacheFailedRequests.push(response);
      return `${FAIL}\n${response.statusText}`;
    }
  } catch (error) {
    console.error(`Failed to get items from list "${listNameCopy}"\n${error.message}`);
    return [];
  }
}

async function createListItemUsingPhysicalToLogicalMapFullResponse(listNameOrGuid, mapFields, item, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  let listNameCopy;
  let urlListItemsBase;
  try {
    if (typeof item !== 'object') {
      throw new Error(`An item object is required!`);
    }
    if (reIsGuidText.test(listNameOrGuid)) {
      listNameCopy = 'having unique identifier: ' + listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    }  else {
      listNameCopy =  listNameOrGuid;
      urlListItemsBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let postUrl = urlListItemsBase + '/items';
    let itemListSafe = convertLogicalToSharePointListFromFieldMap(mapFields, item);
    let postBody = JSON.stringify(itemListSafe);
    let postOptions =  getListOptionsForMethod('POST');
    postOptions.headers.Accept = 'application/json; odata=fullmetadata';
    postOptions.body = postBody;
    postOptions.headers['Content-Length'] = postBody.length;
    postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase);

    let response = await fetch(postUrl, postOptions);
    if (response.status <= 204) {
      return await response.json();
    } else {
      console.error(response);
      cacheFailedRequests.push(response);
      return `${FAIL}\n${response.statusText}`;
    }
  } catch (error) {
    console.error(`Failed to get items from list "${listNameCopy}"\n${error.message}`);
    return [];
  }
}

const getSiteListQueryFromOptions = function (options={itemLimit:5000}, mapFields){
  let {itemLimit, filterClause, selectClause, orderByClause,oDataNextLink} = options;
  if (oDataNextLink) {
    return oDataNextLink.substring(oDataNextLink.indexOf('?') );
  }
  let urlQuery = `?$top=${itemLimit ? itemLimit : 5000}`;
  if (filterClause) {
    urlQuery += '&$filter=' + filterClause;
  }
  if (selectClause) {
    urlQuery += '&$select=' + selectClause;
    if (selectClause.indexOf('Editor/Title') === -1) urlQuery += ',Editor/Title';
    if (selectClause.indexOf('Editor/EMail') === -1) urlQuery += ',Editor/EMail';
    const expandParam = '&$expand=';
    let expandClause = R.compose(
      R.reduce((acc, cur)=> {return acc + (acc.length > expandParam.length ? ',':'') + cur}, expandParam),
      R.uniq,
      R.append('Editor'),
      R.map(s => s.substring(0, s.indexOf('/'))),
      R.filter(s=> s.indexOf('/')>-1),
      R.map(R.trim),
      R.split(',')
    )(selectClause);

    urlQuery += expandClause;
  } else {
    let genericExpandClause = '&$expand=Editor';
    let genericSelect = '&$select=*,Editor/Title,Editor/EMail';
    let aLookupsSingle = fnGetInternalNamesByTypeFromMap ('Lookup', mapFields);
    let aLookupsMulti = fnGetInternalNamesByTypeFromMap ('LookupMulti', mapFields);
    let aLookups = aLookupsSingle.concat(aLookupsMulti);
    if (aLookups.length) {
      let aLookupExpands = R.compose(
        R.uniq,
        R.map(fld => fld.indexOf('/') > - 1 ? fld.substring(0, fld.indexOf('/')): fld)
      )(aLookups);
      genericExpandClause += ','+ aLookupExpands.join(',');
      let aLookupSelects = R.map(fld => fld.indexOf('/') > - 1 ? fld : fld.substring(0, fld.indexOf('/')) +'Title')(aLookups);
      let aLookupMinimalSelects = R.map(key => `${key}/Id,${key}/Title`)(aLookupExpands);
      aLookupSelects = aLookupSelects.concat(aLookupMinimalSelects);
      genericSelect += ',' + R.compose(R.join(','), R.uniq)(aLookupSelects);
    }
    let aUser = fnGetInternalNamesByTypeFromMap ('User', mapFields);
    let aUserMulti = fnGetInternalNamesByTypeFromMap ('UserMulti', mapFields);
    let aAllUserFields = aUser.concat(aUser, aUserMulti);
    if (aAllUserFields.length){
      genericExpandClause += ',' + aAllUserFields.join(',');
      genericSelect +=  ',' + aAllUserFields.map(colName => `${colName}/Id,${colName}/Title,${colName}/EMail`).join(',');
    }
    urlQuery += genericSelect + genericExpandClause;
  }
  if (orderByClause) {
    urlQuery += '&$orderby=' + orderByClause;
  }
  return urlQuery;
};

const getSiteListItems = async function (listNameOrGuid, mapFields, options={itemLimit:5000}, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  let resp;
  let data;
  let url;
  try {
    let urlListBase = '';
    if (reIsGuidText.test(listNameOrGuid)) {
      urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    } else {
      urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    url = `${urlListBase}/items${getSiteListQueryFromOptions(options, mapFields)}`;
    resp = await fetch (url, getListOptionsForMethod('GET'));
    data = await resp.json();
    let aOut;
    if (options.raw) {
      aOut = fGetItemArrayFromRestData(data);
      if (data.hasOwnProperty('odata.nextLink')) aOut.oDataNextLink = data['odata.nextLink'];
    } else {
      aOut =  convertSharePointListToLogicalFromFieldMap(mapFields,  fGetItemArrayFromRestData(data));
      if (data.hasOwnProperty('odata.nextLink')) aOut.oDataNextLink = data['odata.nextLink'];
    }
    return aOut;
  } catch (error) {
    console.error(`Failure getting site list items.\nError: ${error instanceof Error ? error.message : error?.toString()}`);
    console.log(`url = ${url}`);
    console.log(`response status = ${ resp ? resp.status : '-'}`);
    console.log(`response status text= ${ resp ? resp.statusText : '-'}`);
    if (resp) {
      cacheFailedRequests.push(resp);
    }
  }
};

const addAttachmentToListItem = async function (listNameOrGuid, itemId, file, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  try {
    let urlListBase = '';
    if (reIsGuidText.test(listNameOrGuid)) {
      urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    } else {
      urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let url = `${urlListBase}/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
    let postOptions = getListOptionsForMethod('POST');
    postOptions.headers['X-RequestDigest'] = await getRequestDigest(siteBase);
    delete postOptions.headers['Content-Type'];
    postOptions = Object.assign({processData: false, body:file },postOptions);
    let result = await fetch(url,postOptions);
    return await result.json();
  } catch (error) {
    console.error(error);
    return error;
  }
};

const getAttachmentFromListItemNameOrIndex = async function (listNameOrGuid, itemId, fileNameOrIndex, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  try {
    let urlListBase = '';
    if (reIsGuidText.test(listNameOrGuid)) {
      urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    } else {
      urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let url = `${urlListBase}/items(${itemId})/AttachmentFiles`;
    let options = getListOptionsForMethod('GET');
    delete options.headers['Content-Type'];
    if (typeof fileNameOrIndex === 'number'){
      if (Math.floor(fileNameOrIndex) !== fileNameOrIndex || fileNameOrIndex < 0) {
        throw new Error('fileNameOrIndex must be an integer zero or greater');
      }
      let resp = await fetch(url, options);
      let data = await resp.json();
      if (data.value.length === 0) {
        throw new Error(`No attachments found for item ${itemId} in list ${listNameOrGuid}`);
      } else if (fileNameOrIndex >= data.value.length) {
        throw new Error(`Requested attachment index ${fileNameOrIndex} is out of range for item ${itemId} in list ${listNameOrGuid}`);
      } else {
        url += `('${encodeURIComponent(data.value[fileNameOrIndex].FileName)}')/$value`;
        return  fetch(url, options);
      }
    } else if (typeof fileNameOrIndex === 'string'){
      url += `('${encodeURIComponent(fileNameOrIndex)}')/$value`;
      return  fetch(url, options);
    } else {
      throw new Error('fileNameOrIndex must be a number or string');
    }
  } catch (error) {
    console.error(error);
    return error;
  }
};

const deleteAttachmentFromListItemNameOrIndex = async function (listNameOrGuid, itemId, fileNameOrIndex, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  try {
    let urlListBase = '';
    if (reIsGuidText.test(listNameOrGuid)) {
      urlListBase = getSiteListRestUrlByGuid(listNameOrGuid, siteBase);
    } else {
      urlListBase = getSiteListRestUrlByName(listNameOrGuid, siteBase);
    }
    let url = `${urlListBase}/items(${itemId})/AttachmentFiles`;
    let options = getListOptionsForMethod('DELETE');
    delete options.headers['Content-Type'];
    options.headers['X-RequestDigest'] = await getRequestDigest(siteBase);
    options.headers['X-HTTP-Method'] = 'DELETE';
    if (typeof fileNameOrIndex === 'number'){
      if (Math.floor(fileNameOrIndex) !== fileNameOrIndex || fileNameOrIndex < 0) {
        throw new Error('fileNameOrIndex must be an integer zero or greater');
      }
      let resp = await fetch(url, options);
      let data = await resp.json();
      if (data.value.length === 0) {
        throw new Error(`No attachments found for item ${itemId} in list ${listNameOrGuid}`);
      } else if (fileNameOrIndex >= data.value.length) {
        throw new Error(`Requested attachment index ${fileNameOrIndex} is out of range for item ${itemId} in list ${listNameOrGuid}`);
      } else {
        url += `('${encodeURIComponent(data.value[fileNameOrIndex].FileName)}')`;
        return  fetch(url, options);
      }
    } else if (typeof fileNameOrIndex === 'string'){
      url += `('${encodeURIComponent(fileNameOrIndex)}')`;
      return  fetch(url, options);
    } else {
      throw new Error('fileNameOrIndex must be a number or string');
    }
  } catch (error) {
    console.error(error);
    return error;
  }
};

function getLogicalToPhysicalFieldNameForMap (aMapTuple, logicalName) {
  let aPhysicalName = aMapTuple.map(mapPickFirst);
  let aLogicalName = aMapTuple.map(mapPickSecond);
  let idx = aLogicalName.findIndex(R.equals(logicalName));
  if (idx > -1) {
    return aPhysicalName[idx];
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
    let reIsSystemField = /\(system\)/ig;
    this.logicalProperty = logicalProperty;
    this.caption = caption;
    this.editedValue = editedValue;
    this.newValue = newValue;
    this.retain = 'OLD';
    this.isSystem = reIsSystemField.test(caption);
  }
}

function compareMappedEditedToLatestLogicalObjects (mapToUse, edited, latest) {
  let aDifferences = [];
  let aLogicalName = mapToUse.map( mapPickSecond);
  let aCaptions = mapToUse.map(mapPickThird);
  aLogicalName.forEach((logical, idx) =>{
    let editedValue = edited[logical];
    let latestValue = latest[logical];
    let hasDifferences = false;
    if (typeof editedValue !== typeof latestValue) {
      hasDifferences = true;
    } else if (R.not(R.equals(editedValue,latestValue))) {
      hasDifferences = true;
    }
    if (hasDifferences) {
      let differenceSpecification = new ObjectVersionDifference(logical, aCaptions[idx], editedValue, latestValue);
      if (differenceSpecification.isSystem) {
        differenceSpecification.retain = 'NEW';
      }
      aDifferences.push(differenceSpecification);
    }
  });
  return aDifferences;
}

const getCurrentUserInSite =  async function (siteBase = SITE_BASE) {
  if (!cacheCurrentUserInSite[siteBase]) {
    let resp = await fetch(`${siteBase}/_api/web/currentUser`, getListOptionsForMethod('GET'));
    let data = await resp.json();
    cacheCurrentUserInSite[siteBase] = data;
  }
  return cacheCurrentUserInSite[siteBase];
};

const getGroupAndMembersByNameOrIdInSite  = async function (nameOrId, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  const urlGroup = `${siteBase}/_api/web/SiteGroups/GetBy${typeof nameOrId === 'number' ? `ID(${nameOrId})` : `Name('${encodeURIComponent(nameOrId)}')`}`;
  const urlMembers = urlGroup + '/users';
  const options = getListOptionsForMethod('GET');
  try {
    const [oGroup, oUserQuery] = await Promise.all([
      fetch (urlGroup, options).then (r => r.json()),
      fetch (urlMembers, options).then (r => r.json()),
    ]);
    return R.assoc('members', R.pathOr([],['value'], oUserQuery), oGroup);
  } catch (ex) {
    return {message: ex instanceof Error ? ex.message : ex?.toString()};
  }
};

function getSemVer() {
  let parts = SemVer.split('.');
  return {major: parts[0], minor: parts[1], revision: parts[2]};
}

function popLastFailedRequest() {
  return cacheFailedRequests.pop();
}
function getCountFailedResponse() {
  return cacheFailedRequests.length;
}

// Export all public API as named exports
export {
  SemVer,
  getSiteListItems,
  updateOrCreateListItemUsingPhysicalToLogicalMap,
  createListItemUsingPhysicalToLogicalMapFullResponse,
  deleteItemInListWithId,
  getLogicalToPhysicalFieldNameForMap,
  getRequestDigest,
  createTemplateRecordFromMapAndPassedDefaultKeySet,
  fConvertIsoDateTextToDate as convertIsoDateTextToDate,
  localDateToSharePoint,
  compareMappedEditedToLatestLogicalObjects,
  getCurrentUserInSite,
  getGroupAndMembersByNameOrIdInSite,
  popLastFailedRequest,
  getCountFailedResponse,
  convertLogicalToSharePointListFromFieldMap,
  convertSharePointListToLogicalFromFieldMap,
  getContextInfoForSite,
  getSemVer,
  getSiteBase,
  setSiteBase,
  addAttachmentToListItem,
  getAttachmentFromListItemNameOrIndex,
  deleteAttachmentFromListItemNameOrIndex
};