let csrf='';
export function setCsrf(value){csrf=value||'';}
export async function api(path,{method='GET',body}={}){
  const response=await fetch('/api'+path,{method,credentials:'same-origin',headers:{...(body?{'Content-Type':'application/json'}:{}),...(method!=='GET'?{'X-CSRF-Token':csrf}: {})},body:body?JSON.stringify(body):undefined});
  if(response.status===204)return null;
  const data=await response.json().catch(()=>({}));
  if(!response.ok){const error=new Error(data.message||'Request could not be completed.');error.code=data.error;error.status=response.status;throw error;}
  return data;
}
