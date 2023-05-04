/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

// see https://sourcegraph.com/github.com/aspnet/AspNetCore@bd65275148abc9b07a3b59797a88d485341152bf/-/blob/src/Components/Web.JS/src/Boot.Server.ts#L41:9
export { };
declare global {
    interface Window {
        lancerchaine: any;
    }  
}
try {
    var exports = {};
    window.lancerchaine = async (event: Office.AddinCommands.Event) => {
        console.log("lancerchaine");
        await callStaticLocalComponentMethod();
        console.log("nod");
        event.completed();
    }
    Office.actions.associate("lancerchaine", (window).lancerchaine);
}
catch (err){
    console.log("erreur windows : " + err.message);
}

 async function callStaticLocalComponentMethod() {
    //window.dispatchEvent(new Event('myEvent'));
    console.log("avant");
    try {               
        var dato = "init";
        dato =await DotNet.invokeMethodAsync("BlazorAddIn", "Localfunction");                                       
        console.log("fin demarrage : " + dato);
    }
    catch (err) {        
        console.log("erreur : " + err.message);      
    }
    finally {       
        console.log("après");
    }
}
