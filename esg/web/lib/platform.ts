type NativeReply = {id:string;ok:boolean;cancelled?:boolean;error?:string};
declare global {
  interface Window {
    ESGNative?: {postMessage(message:string):void;onmessage:((event:{data:string})=>void)|null};
    esgAndroidBack?:()=>boolean;
    esgHasUnsavedChanges?:boolean;
  }
}
const pending=new Map<string,{resolve:(reply:NativeReply)=>void;timer:ReturnType<typeof setTimeout>}>();
function notice(message:string,error=false){window.dispatchEvent(new CustomEvent('esg:platform-notice',{detail:{message,error}}))}
function nativeRequest(payload:Record<string,unknown>):Promise<NativeReply>{
  const bridge=window.ESGNative;
  if(!bridge)return Promise.resolve({id:'',ok:false,error:'Android integration unavailable.'});
  bridge.onmessage=(event)=>{try{
    const reply=JSON.parse(event.data);
    if(!reply||typeof reply!=='object'||typeof reply.id!=='string')return;
    const request=pending.get(reply.id);
    if(request){
      clearTimeout(request.timer);pending.delete(reply.id);
      const valid=typeof reply.ok==='boolean'&&(reply.cancelled===undefined||typeof reply.cancelled==='boolean')&&(reply.error===undefined||typeof reply.error==='string')&&!(reply.ok&&reply.cancelled);
      request.resolve(valid?reply:{id:reply.id,ok:false,error:'Android returned an invalid response.'});
    }
  }catch{notice('Android returned an unreadable response.',true)}};
  const id=crypto.randomUUID();
  const message=JSON.stringify({id,...payload});
  if(message.length>5*1024*1024*2+8192)return Promise.resolve({id,ok:false,error:'This export is too large for the Android document bridge.'});
  return new Promise(resolve=>{
    const timer=setTimeout(()=>{pending.delete(id);resolve({id,ok:false,error:'No response from Android. Check whether a file dialog is still open.'})},300000);
    pending.set(id,{resolve,timer});
    try{bridge.postMessage(message)}catch{clearTimeout(timer);pending.delete(id);resolve({id,ok:false,error:'Unable to open the Android action.'})}
  });
}
export async function exportDocument(name:string,text:string,mime='application/json'){
  if(window.ESGNative){
    if(!['application/json','text/csv','text/plain'].includes(mime)){notice('This document format is not supported on Android.',true);return}
    if(new TextEncoder().encode(text).byteLength>5*1024*1024){notice('Export must be 5 MiB or smaller.',true);return}
    const reply=await nativeRequest({action:'export',name,mime,text});
    if(reply.cancelled)notice('Export cancelled');
    else if(!reply.ok)notice(reply.error||'Unable to export the file.',true);
    else notice('File exported');
    return;
  }
  const url=URL.createObjectURL(new Blob([text],{type:mime}));
  const a=document.createElement('a');a.href=url;a.download=name;a.click();
  setTimeout(()=>URL.revokeObjectURL(url),2000);
}
export async function printReport(){
  if(window.ESGNative){
    // Two frames let a just-selected report finish layout before Android captures it.
    await new Promise<void>(resolve=>requestAnimationFrame(()=>requestAnimationFrame(()=>resolve())));
    const reply=await nativeRequest({action:'print'});
    if(!reply.ok&&!reply.cancelled)notice(reply.error||'Unable to open printing.',true);
  }else window.print();
}
