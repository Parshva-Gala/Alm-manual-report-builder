import React from 'react';
import {createRoot} from 'react-dom/client';
import Workbench from './components/workbench';
import './styles.css';
class ErrorBoundary extends React.Component<{children:React.ReactNode},{error:boolean}>{state={error:false};static getDerivedStateFromError(){return{error:true}}render(){return this.state.error?<main style={{padding:32,maxWidth:650,margin:'auto'}}><h1>Unable to open this view</h1><p style={{margin:'20px 0'}}>Your saved workspace has not been changed. Close and reopen the application, or update Android System WebView.</p><button className="primary-button" onClick={()=>location.reload()}>Reopen workspace</button></main>:this.props.children}}
createRoot(document.getElementById('root')!).render(<ErrorBoundary><Workbench/></ErrorBoundary>);
