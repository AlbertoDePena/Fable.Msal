namespace Fable.Msal

open Fable.Core
open Fable.Core.JS
open Fable.Core.JsInterop

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
    
    let private ofChoice x =
        match x with
        | Choice1Of2 o -> Ok o
        | Choice2Of2 e -> Error e

    let private map f (computation: Async<'t>) = async {
        let! x = computation
        return f x
    }

    let private getUserNameInterop () : string = import "getUserName" "./Msal.js"

    let private getProfileInterop () : GraphUserInfo Promise = import "getProfile" "./Msal.js"

    let private getMailInterop () : GraphMailInfo Promise = import "getMail" "./Msal.js"

    let signIn (config : MsalConfig) : unit = import "signIn" "./Msal.js"

    let signOut () : unit = import "signOut" "./Msal.js"

    let getUserName () : string option =
        let userName = getUserNameInterop ()
        if System.String.IsNullOrWhiteSpace userName
        then None
        else Some userName
        
    let getProfile () : Async<Result<GraphUserInfo, exn>> =
        getProfileInterop ()
        |> Async.AwaitPromise
        |> Async.Catch
        |> map ofChoice

    let getMail () : Async<Result<GraphMailInfo, exn>> =
        getMailInterop ()
        |> Async.AwaitPromise
        |> Async.Catch
        |> map ofChoice