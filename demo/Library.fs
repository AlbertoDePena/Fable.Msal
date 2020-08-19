module Demo 

    open Fable.Msal
    open Browser.Dom

    let map f (computation: Async<'t>) = async {
        let! x = computation
        return f x
    }

    Msal.signIn ({
        clientId = "0c89962d-0574-46e7-8644-656286b8c0eb"
        authority = "https://login.microsoftonline.com/dfe6522a-e1ef-4132-a50b-afa26c14bc41"
        redirectUri = "http://localhost:8080" 
        cacheLocation = "sessionStorage"
        storeAuthStateInCookie = true
        useLoginRedirect = true
    })

    let signOutBtn = document.getElementById "signOutBtn"
    let loadProfileBtn = document.getElementById "loadProfileBtn"
    let loadMailBtn = document.getElementById "loadMailBtn"
    let getUserNameBtn = document.getElementById "getUserNameBtn"

    signOutBtn.addEventListener("click", fun _ -> Msal.signOut())

    loadProfileBtn.addEventListener("click", 
        fun _ -> Msal.getProfile() |> map (fun profile -> console.log(profile)) |> Async.StartImmediate)

    loadMailBtn.addEventListener("click", 
        fun _ -> Msal.getMail() |> map (fun mail -> console.log(mail)) |> Async.StartImmediate)

    getUserNameBtn.addEventListener("click", 
        fun _ -> 
            let userName = Msal.getUserName() |> Option.defaultValue "N/A"
            console.log (userName))
