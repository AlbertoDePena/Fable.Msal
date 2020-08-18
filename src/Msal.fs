namespace Fable.Msal

open Fable.Core
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

type MsalAuth = { 
    clientId : string
    authority : string
    redirectUri: string }

type MsalConfiguration = {
    auth : MsalAuth
}    

type IMsal =
    abstract SignInPopup : MsalConfiguration -> unit
    abstract SignInRedirect : MsalConfiguration -> unit
    abstract SignOut : unit -> unit
    abstract GetUserName : unit -> string
    abstract GetProfile : unit -> GraphUserInfo Promise
    abstract GetMail : unit -> GraphMailInfo Promise

[<AutoOpen>]
module Core =
    
    [<ImportAll("./Msal.js")>]
    let Msal: IMsal = jsNative