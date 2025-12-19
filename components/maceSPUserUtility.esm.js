import * as R from 'https://cdn.jsdelivr.net/npm/ramda@0.32.0/+esm'



const SemVer = '2.1.2';

let SITE_BASE;
if (document?.currentScript) {
  let src = document.currentScript.src;
  let reFindSite = /(.+)\/siteAssets\/(.*)/i;
  if (reFindSite.test(src)){
    let match = reFindSite.exec(src);
    console.log(`maceSPUserUtility dynamically determined its site base to be ${match[1]}`);
    SITE_BASE = match[1];
  }
}

let cacheCurrentUserInSite = {};
const OK = 'OK';
const FAIL = 'FAIL';

const fGetItemArrayFromRestData = R.pathOr([],['value']);

const getGroupApiRelativeUrlFromNameOrId = (nameOrId) =>
  `_api/web/SiteGroups/GetBy${typeof nameOrId === 'number' ? `ID(${nameOrId})` : `Name('${encodeURIComponent(nameOrId)}')`}`;

function autoPrefixSiteUrlWithBase(siteBase) {
  if (typeof siteBase === 'string' && !siteBase.startsWith('https://')) {
    return `${SITE_BASE}${ siteBase.startsWith('/') ? '' : '/'}${siteBase}`;
  } else {
    return siteBase;
  }
}

function getSiteBase() { return SITE_BASE; }
function setSiteBase(url) { SITE_BASE = url; }

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
}

const getCurrentUserInSite = async function (siteBase = SITE_BASE) {
  if (!cacheCurrentUserInSite[siteBase]) {
    let resp = await fetch(`${siteBase}/_api/web/currentUser`, getListOptionsForMethod('GET'));
    let data = await resp.json();
    cacheCurrentUserInSite[siteBase] = data;
  }
  return cacheCurrentUserInSite[siteBase];
};

const getGroupAndMembersByNameOrIdInSite = async function (nameOrId, siteBase = SITE_BASE) {
  siteBase = autoPrefixSiteUrlWithBase(siteBase);
  const urlGroup = `${siteBase}/${getGroupApiRelativeUrlFromNameOrId(nameOrId)}`;
  const urlMembers = urlGroup + '/users';
  const options = getListOptionsForMethod('GET');
  try {
    const [oGroup, oUserQuery] = await Promise.all([
      fetch(urlGroup, options).then(r => r.json()),
      fetch(urlMembers, options).then(r => r.json()),
    ]);
    return R.assoc('members', R.pathOr([],['value'], oUserQuery), oGroup);
  } catch (ex) {
    return {message: ex instanceof Error ? ex.message : ex?.toString()};
  }
};

const getAllSiteGroups = async function (siteBase = SITE_BASE) {
  let url = `${siteBase}/_api/web/SiteGroups?$select=Id,Title,Description,OwnerTitle`;
  let aGroups = await fetch(url, getListOptionsForMethod('GET'))
    .then(r=> r.json())
    .then(d=> fGetItemArrayFromRestData(d));
  return aGroups;
};

const getGroupMembersByGroupNameOrId = async function (nameOrId, siteBase = SITE_BASE) {
  let url = `${siteBase}/${getGroupApiRelativeUrlFromNameOrId(nameOrId)}/users`;
  let aMembers = fGetItemArrayFromRestData(await fetch(url, getListOptionsForMethod('GET')).then(r=>r.json()));
  return aMembers;
};

const getAllSiteMembers = async function (usersGroupsOrBoth = 'U', siteBase = SITE_BASE) {
  let aGroups = await getAllSiteGroups(siteBase);
  let aGroupMemberSets = await Promise.all(aGroups.map(grp => getGroupMembersByGroupNameOrId(grp.Id, siteBase)));
  let fnFilter;
  switch (usersGroupsOrBoth.substring(0,1).toUpperCase()) {
    case 'U':
      fnFilter = (obj) => obj?.UserId !== null;
      break;
    case 'G':
      fnFilter = (obj) => obj?.UserId === null;
      break;
    default:
      fnFilter = R.T;
  }
  let aMapPeople = R.compose(
    R.sortBy(R.view(R.lensProp('Title'))),
    R.uniqBy(R.view(R.lensProp('Id'))),
    R.filter(fnFilter),
    R.flatten
  )(aGroupMemberSets);
  return aMapPeople;
};

const getAllSiteUsers = async function (filters={email:'', title:''}, siteBase = SITE_BASE) {
  let url = `${siteBase}/_api/web/SiteUsers?$select=Id,Title,Email`;
  let resp = await fetch(url, getListOptionsForMethod('GET'));
  let data = await resp.json();
  let aUsers = R.pathOr([],['value'], data);
  if (filters?.email) {
    let reEmail = new RegExp(filters.email,"i");
    aUsers = aUsers.filter(usr => reEmail.test(usr.Email));
  }
  if (filters?.title) {
    let reName = new RegExp(filters.title,"i");
    aUsers = aUsers.filter(usr => reName.test(usr.Title));
  }
  return aUsers;
};

export {
  SemVer,
  getSiteBase,
  setSiteBase,
  getCurrentUserInSite,
  getGroupAndMembersByNameOrIdInSite,
  getGroupMembersByGroupNameOrId,
  getAllSiteGroups,
  getAllSiteMembers,
  getAllSiteUsers,
  OK,
  FAIL
};