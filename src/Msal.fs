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
module internal Interop =

    let signIn (config : MsalConfig) : unit = import "signIn" "./Msal.js"

    let signOut () : unit = import "signOut" "./Msal.js"

    let getUserName () : string = import "getUserName" "./Msal.js"

    let getProfile () : GraphUserInfo Promise = import "getProfile" "./Msal.js"

    let getMail () : GraphMailInfo Promise = import "getMail" "./Msal.js"

[<RequireQualifiedAccess>]
module Msal =
    
    let private toAsyncResult (computation: Async<'t>) = async {
        let! choice = computation |> Async.Catch
        
        let result =
            match choice with
            | Choice1Of2 o -> Ok o
            | Choice2Of2 e -> Error e
        
        return result
    }

    /// Trigger sign in Msal flow
    let signIn (config : MsalConfig) : unit = 
        Interop.signIn config

    /// Trigger sign out Msal flow
    let signOut () : unit = 
        Interop.signOut ()

    /// Get user name from Msal bearer token claims
    let getUserName () : string option =
        let userName = Interop.getUserName ()
        if System.String.IsNullOrWhiteSpace userName
        then None
        else Some userName
        
    /// Get user profile from graph API
    let getProfile () : Async<Result<GraphUserInfo, exn>> =
        Interop.getProfile()
        |> Async.AwaitPromise
        |> toAsyncResult

    /// Get user mail from graph API
    let getMail () : Async<Result<GraphMailInfo, exn>> =
        Interop.getMail ()
        |> Async.AwaitPromise
        |> toAsyncResult