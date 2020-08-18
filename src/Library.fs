namespace Fable.Msal

open Fable.Core
open Fable.Core.JS
open Browser.Dom

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

module Test =

    Msal.SignInRedirect ({
        auth = { 
            clientId = "0c89962d-0574-46e7-8644-656286b8c0eb"
            authority = "https://login.microsoftonline.com/dfe6522a-e1ef-4132-a50b-afa26c14bc41"
            redirectUri = "http://localhost:8080" }
    })

    let signOffBtn = document.getElementById "signOffBtn"
    let loadProfileBtn = document.getElementById "loadProfileBtn"
    let loadMailBtn = document.getElementById "loadMailBtn"
    let getUserNameBtn = document.getElementById "getUserNameBtn"

    signOffBtn.addEventListener("click", fun _ -> Msal.SignOut())

    loadProfileBtn.addEventListener("click", 
        fun _ -> Msal.GetProfile() |> ignore)

    loadMailBtn.addEventListener("click", 
        fun _ -> Msal.GetMail() |> ignore)

    getUserNameBtn.addEventListener("click", 
        fun _ -> console.log (Msal.GetUserName()))