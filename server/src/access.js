import { check } from './errors.js';
import { rbacBuiltInRoleSeeds_,rbacNormalizeRole_,iamParseDepartments_,iamDepartmentKey_ } from './source-rules.js';
export const USER_HEADERS=['Email','Display Name','Role','Active','Updated By','Updated At','Departments'];
export const ROLE_HEADERS=['Role Name','Description','Built In','Active','Scope Type','Allowed Departments','Allowed Users','Permissions JSON','Updated By','Updated At'];
export const GRANT_HEADERS=['Grant ID','Request ID','User Email','Original Role','Original Departments JSON','Original Active','Granted Role','Granted Departments JSON','Start At','End At','Status','Created At','Activated At','Reverted At','Last Error'];
export const AUTH_TABLES=['User_Access','Role_Definitions','Temporary_Access_Grants','Dashboard_Credentials'];
export const truth=v=>v===true||String(v).toLowerCase()==='true';
export const text=v=>String(v??'').trim();
export function requireSchema(table,headers,name,{optional=false}={}){
  if(table===null&&optional)return;
  check(Array.isArray(table)&&table.length>0&&headers.every((h,i)=>table[0][i]===h),503,'SCHEMA_MISMATCH',`${name} has missing or unexpected columns.`);
}
export function getAccess(tables,email,domains,now=Date.now(),storage={}){
  email=text(email).toLowerCase();
  check(domains.includes(email.split('@')[1]),403,'ACCESS_DENIED','Use an approved company Google account.');
  requireSchema(tables.User_Access,USER_HEADERS.slice(0,6),'User_Access');
  const users=tables.User_Access.slice(1).filter(r=>text(r[0]).toLowerCase()===email);
  check(users.length===1&&truth(users[0][3]),403,'ACCESS_DENIED','No unique active dashboard account was found.');
  const user=users[0];
  requireSchema(tables.Temporary_Access_Grants,GRANT_HEADERS,'Temporary_Access_Grants',{optional:true});
  const unresolved=(tables.Temporary_Access_Grants||[]).slice(1).filter(r=>text(r[2]).toLowerCase()===email&&!['REVERTED','CANCELLED'].includes(r[10]));
  check(!unresolved.some(r=>['ERROR','APPLYING'].includes(r[10])||!Number.isFinite(Date.parse(r[9]))||Date.parse(r[9])<=now),403,'ACCESS_PAUSED','A temporary grant expired or needs recovery. Use the existing dashboard to restore it.');
  let roles;
  if(tables.Role_Definitions===null)roles=rbacBuiltInRoleSeeds_().map(rbacNormalizeRole_);
  else{
    requireSchema(tables.Role_Definitions,ROLE_HEADERS,'Role_Definitions');
    roles=tables.Role_Definitions.slice(1).map(r=>rbacNormalizeRole_({name:r[0],description:r[1],builtIn:r[2],active:r[3],scopeType:r[4],allowedDepartments:r[5],allowedUsers:r[6],permissions:r[7]}));
  }
  const found=roles.filter(r=>r.name.toLowerCase()===text(user[2]).toLowerCase());
  check(found.length===1&&found[0].active,403,'ACCESS_DENIED','The assigned role is missing, duplicated or inactive.');
  const role=found[0],departments=iamParseDepartments_(user[6]);
  const departmentScoped=role.scopeType==='DEPARTMENTS';
  check(!departmentScoped||(departments.length>0&&!departments.some(d=>/^(all|\*)$/i.test(d))),403,'ACCESS_DENIED','Assign specific departments before granting access.');
  if(departmentScoped&&storage.mode==='google')check(Boolean(storage.iamStorageOwner&&storage.serviceAccount&&storage.iamStorageOwner===storage.serviceAccount),503,'STORAGE_OWNER_MISMATCH','Department access is paused until the Cloud Run service identity matches the approved Sheets storage owner. Ask IT to check staging configuration.');
  // Respect explicit role restrictions. Legacy code did not consistently enforce
  // USERS scope; this port denies access outside the configured allowlist.
  check(role.scopeType!=='USERS'||role.allowedUsers.includes(email),403,'ACCESS_DENIED','This role is restricted to its configured users.');
  check(!departmentScoped||!role.allowedDepartments.length||departments.every(d=>role.allowedDepartments.some(a=>iamDepartmentKey_(a)===iamDepartmentKey_(d))),403,'ACCESS_DENIED','Department assignment is outside the role allowlist.');
  return {email,name:text(user[1]),role:role.name,permissions:role.permissions,departments,departmentScoped,isSystemAdmin:role.builtIn&&role.name==='Admin'};
}
export function requirePermission(access,key){check(access.permissions[key]===true,403,'ACCESS_DENIED',`Your role does not permit ${key}.`);}
