import { GoogleAuth } from 'google-auth-library';
import { AppError } from './errors.js';

export class SheetsRepository {
  constructor(spreadsheetId){
    if(!spreadsheetId)throw Error('GOOGLE_SHEETS_ID is required in google mode.');
    this.spreadsheetId=spreadsheetId;
    this.auth=new GoogleAuth({scopes:['https://www.googleapis.com/auth/spreadsheets']});
  }
  async request(suffix,options={}){
    const client=await this.auth.getClient();
    try{return (await client.request({url:`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(this.spreadsheetId)}${suffix}`,...options})).data;}
    catch(error){console.error('Sheets API failed',{status:error.response?.status});throw new AppError(503,'SHEETS_UNAVAILABLE','Unable to access the configured spreadsheet. Ask IT to check service account access.');}
  }
  async tables(names){
    // Batch related reads; never accept a client-provided sheet name or range.
    const meta=await this.request('?fields=sheets.properties.title');
    const available=new Set(meta.sheets.map(s=>s.properties.title));
    const present=names.filter(n=>available.has(n));
    const out=Object.fromEntries(names.map(n=>[n,null]));
    if(!present.length)return out;
    const query=new URLSearchParams({valueRenderOption:'FORMATTED_VALUE'});
    present.forEach(n=>query.append('ranges',`'${n.replaceAll("'","''")}'!A:EZ`));
    const data=await this.request('/values:batchGet?'+query);
    present.forEach((name,i)=>out[name]=data.valueRanges[i]?.values||[]);
    return out;
  }
  async credentialResult(row,attempts,lockedUntil,lastLoginAt){
    await this.request(`/values/${encodeURIComponent(`'Dashboard_Credentials'!H${row}:J${row}`)}?valueInputOption=RAW`,{
      method:'PUT',data:{values:[[attempts,lockedUntil,lastLoginAt]]}
    });
  }
}

export class MemoryRepository {
  constructor(data){this.data=structuredClone(data);}
  async tables(names){return Object.fromEntries(names.map(n=>[n,structuredClone(this.data[n]??null)]));}
  async credentialResult(row,attempts,lockedUntil,lastLoginAt){
    const values=this.data.Dashboard_Credentials[row-1];
    values[7]=attempts;values[8]=lockedUntil;values[9]=lastLoginAt;
  }
}
