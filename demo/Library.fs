module Demo 

    open Fable.Msal
    open Browser.Dom

    Msal.signIn ({
        clientId = "0c89962d-0574-46e7-8644-656286b8c0eb"
        authority = "https://login.microsoftonline.com/dfe6522a-e1ef-4132-a50b-afa26c14bc41"
        redirectUri = "http://localhost:8080" 
        cacheLocation = "sessionStorage"
        storeAuthStateInCookie = true
        useLoginRedirect = true
    })

    let getTokenBtn = document.getElementById "getTokenBtn"
    let signOutBtn = document.getElementById "signOutBtn"
    let loadProfileBtn = document.getElementById "loadProfileBtn"
    let loadMailBtn = document.getElementById "loadMailBtn"
    let getUserNameBtn = document.getElementById "getUserNameBtn"

    getTokenBtn.addEventListener("click", 
        fun _ -> 
            let computation = async {
                let! tokenResult = Msal.getToken ()

                match tokenResult with
                | Ok token -> console.log(token)
                | Error error -> console.error(error)
            }

            computation |> Async.StartImmediate)

    signOutBtn.addEventListener("click", fun _ -> Msal.signOut())

    loadProfileBtn.addEventListener("click", 
        fun _ -> 
            let computation = async {
                let! profileResult = Msal.getProfile ()

                match profileResult with
                | Ok profile -> console.log(profile)
                | Error error -> console.error(error)
            }

            computation |> Async.StartImmediate)

    loadMailBtn.addEventListener("click", 
        fun _ -> 
            let computation = async {
                let! mailResult = Msal.getMail ()

                match mailResult with
                | Ok mail -> console.log(mail)
                | Error error -> console.error(error)
            }
            
            computation |> Async.StartImmediate)

    getUserNameBtn.addEventListener("click", 
        fun _ -> 
            let userName = Msal.getUserName() |> Option.defaultValue "N/A"
            console.log (userName))
