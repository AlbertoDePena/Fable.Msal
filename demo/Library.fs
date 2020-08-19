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

    let signOffBtn = document.getElementById "signOffBtn"
    let loadProfileBtn = document.getElementById "loadProfileBtn"
    let loadMailBtn = document.getElementById "loadMailBtn"
    let getUserNameBtn = document.getElementById "getUserNameBtn"

    signOffBtn.addEventListener("click", fun _ -> Msal.signOut())

    loadProfileBtn.addEventListener("click", 
        fun _ -> Msal.getProfile() |> Promise.iter (fun profile -> console.log(profile)))

    loadMailBtn.addEventListener("click", 
        fun _ -> Msal.getMail() |> Promise.iter (fun mail -> console.log(mail)))

    getUserNameBtn.addEventListener("click", 
        fun _ -> console.log (Msal.getUserName()))
