namespace Fable.Msal

open Fable.Core.JsInterop
open Fable.Core.JS

type GraphEmailAddress = { address : string }

type GraphFromEmailAddress = { emailAddress : GraphEmailAddress }

type GraphMailItem = {
    from : GraphFromEmailAddress
    subject : string
    bodyPreview : string }

type GraphMailInfo = { value : GraphMailItem array }

type GraphUserInfo = {
    businessPhones : string array
    displayName : string
    givenName : string
    id : string
    jobTitle : string
    mail : string
    mobilePhone : string
    officeLocation : string
    preferredLanguage : string
    surname : string
    userPrincipalName : string }

type MsalConfig = {
    clientId : string
    authority : string
    redirectUri: string
    cacheLocation : string
    storeAuthStateInCookie : bool
    useLoginRedirect : bool }

[<RequireQualifiedAccess>]
module Msal =
    
    let signIn (config : MsalConfig) : unit = import "signIn" "./Msal.js"

    let signOut () : unit = import "signOut" "./Msal.js"

    let getUserName () : string = import "getUserName" "./Msal.js"

    let getProfile () : GraphUserInfo Promise = import "getProfile" "./Msal.js"

    let getMail () : GraphMailInfo Promise = import "getMail" "./Msal.js"