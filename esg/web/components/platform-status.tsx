import {useEffect,useState} from 'react';
export function PlatformStatus(){
 const [status,setStatus]=useState<{message:string;error:boolean}|null>(null);
 useEffect(()=>{const listener=(event:Event)=>setStatus((event as CustomEvent).detail);window.addEventListener('esg:platform-notice',listener);return()=>window.removeEventListener('esg:platform-notice',listener)},[]);
 useEffect(()=>{if(!status)return;const timer=setTimeout(()=>setStatus(null),6500);return()=>clearTimeout(timer)},[status]);
 return status?<div className="native-notice" data-tone={status.error?'error':'success'} role={status.error?'alert':'status'}>{status.message}</div>:null;
}
