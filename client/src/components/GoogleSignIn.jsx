import {useEffect,useRef} from 'react';
export default function GoogleSignIn({clientId,onCredential}){
  const host=useRef(null);
  useEffect(()=>{
    if(!clientId)return;
    let timer,disposed=false;
    function mount(){
      if(disposed||!window.google?.accounts?.id||!host.current)return false;
      host.current.replaceChildren();
      window.google.accounts.id.initialize({client_id:clientId,callback:r=>r.credential&&onCredential(r.credential),auto_select:false,cancel_on_tap_outside:true});
      window.google.accounts.id.renderButton(host.current,{theme:'outline',size:'large',shape:'pill',text:'signin_with',width:290,logo_alignment:'left'});
      return true;
    }
    const script=document.querySelector('script[data-dc-gsi]')||document.createElement('script');
    if(!script.dataset.dcGsi){script.src='https://accounts.google.com/gsi/client';script.async=true;script.defer=true;script.dataset.dcGsi='1';document.head.append(script);}
    script.addEventListener('load',mount);let tries=0;
    const poll=()=>{if(mount()||++tries>100)return;timer=setTimeout(poll,100);};poll();
    return()=>{disposed=true;clearTimeout(timer);script.removeEventListener('load',mount);};
  },[clientId,onCredential]);
  return <div className="google-signin" ref={host} aria-label="Continue with Google"/>;
}
