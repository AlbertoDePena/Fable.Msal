import {
    PublicClientApplication, AuthorizationUrlRequest, SilentRequest,
    AuthenticationResult, Configuration, LogLevel, AccountInfo, InteractionRequiredAuthError,
    EndSessionRequest, RedirectRequest, PopupRequest
} from "@azure/msal-browser/dist";

import { isIE } from "./Util";

/**
 * Configuration class for @azure/msal-browser: 
 * https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-browser/modules/_src_config_configuration_.html
 * @type {Configuration}
 */
const MSAL_CONFIG = {
    auth: {
        clientId: "",
        authority: "",
        redirectUri: "" 
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: isIE,
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) { return; }
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                }
            }
        }
    }
};

/**
 * AuthModule for application - handles authentication in app.
 */
export class AuthModule {

    // private myMSALObj: PublicClientApplication; // https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-browser/classes/_src_app_publicclientapplication_.publicclientapplication.html
    // private account: AccountInfo; // https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-common/modules/_src_account_accountinfo_.html
    // private loginRedirectRequest: RedirectRequest; // https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-browser/modules/_src_request_redirectrequest_.html
    // private loginRequest: PopupRequest; // https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-browser/modules/_src_request_popuprequest_.html
    // private profileRedirectRequest: RedirectRequest;
    // private profileRequest: PopupRequest;
    // private mailRedirectRequest: RedirectRequest;
    // private mailRequest: PopupRequest;
    // private silentProfileRequest: SilentRequest; // https://azuread.github.io/microsoft-authentication-library-for-js/ref/msal-browser/modules/_src_request_silentrequest_.html
    // private silentMailRequest: SilentRequest;

    constructor(config) {
        MSAL_CONFIG.auth.clientId = config.clientId;
        MSAL_CONFIG.auth.authority = config.authority;
        MSAL_CONFIG.auth.redirectUri = config.redirectUri;
        MSAL_CONFIG.cache.cacheLocation = config.cacheLocation;
        MSAL_CONFIG.cache.storeAuthStateInCookie = config.storeAuthStateInCookie;

        this.myMSALObj = new PublicClientApplication(MSAL_CONFIG);
        this.account = null;
        this.setRequestObjects();
    }

    /**
     * Initialize request objects used by this AuthModule.
     */
    setRequestObjects() {
        this.loginRequest = {
            scopes: []
        };

        this.loginRedirectRequest = {
            ...this.loginRequest,
            redirectStartPage: window.location.href
        };

        this.profileRequest = {
            scopes: ["User.Read"]
        };

        this.profileRedirectRequest = {
            ...this.profileRequest,
            redirectStartPage: window.location.href
        };

        // Add here scopes for access token to be used at MS Graph API endpoints.
        this.mailRequest = {
            scopes: ["Mail.Read"]
        };

        this.mailRedirectRequest = {
            ...this.mailRequest,
            redirectStartPage: window.location.href
        };

        this.silentProfileRequest = {
            scopes: ["openid", "profile", "User.Read"],
            account: null,
            forceRefresh: false
        };

        this.silentMailRequest = {
            scopes: ["openid", "profile", "Mail.Read"],
            account: null,
            forceRefresh: false
        };
    }

    /**
     * Calls getAllAccounts and determines the correct account to sign into, currently defaults to first account found in cache.
     * TODO: Add account chooser code
     * 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    getAccount() {
        // need to call getAccount here?
        const currentAccounts = this.myMSALObj.getAllAccounts();
        if (currentAccounts === null) {
            console.log("No accounts detected");
            return null;
        }

        if (currentAccounts.length > 1) {
            // Add choose account code here
            console.log("Multiple accounts detected, need to add choose account code.");
            return currentAccounts[0];
        } else if (currentAccounts.length === 1) {
            return currentAccounts[0];
        }
    }

    /**
     * Checks whether we are in the middle of a redirect and handles state accordingly. Only required for redirect flows.
     * 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/initialization.md#redirect-apis
     */
    loadAuthModule() {
        const that = this;
        return that.myMSALObj.handleRedirectPromise().then((resp) => {
            return that.handleResponse(resp);
        }).catch(console.error);
    }

    /**
     * Handles the response from a popup or redirect. If response is null, will check if we have any accounts and attempt to sign in.
     * @param response 
     */
    handleResponse(response) {
        if (response !== null) {
            this.account = response.account;
        } else {
            this.account = this.getAccount();
        }

        return this.account;
    }

    /**
     * Calls loginRedirect.
     */
    loginRedirect() {
        this.myMSALObj.loginRedirect(this.loginRedirectRequest);
    }

    /**
     * Calls loginPopup.
     */
    loginPopup() {
        this.myMSALObj.loginPopup(this.loginRequest).then((resp) => {
            this.handleResponse(resp);
        }).catch(console.error);
    }

    /**
     * Logs out of current account.
     */
    logout() {
        const logOutRequest = {
            account: this.account
        };

        this.myMSALObj.logout(logOutRequest);
    }

    /**
     * Gets the token to read user profile data from MS Graph silently, or falls back to interactive redirect.
     */
    async getProfileTokenRedirect() {
        this.silentProfileRequest.account = this.account;
        return this.getTokenRedirect(this.silentProfileRequest, this.profileRedirectRequest);
    }

    /**
     * Gets the token to read user profile data from MS Graph silently, or falls back to interactive popup.
     */
    async getProfileTokenPopup() {
        this.silentProfileRequest.account = this.account;
        return this.getTokenPopup(this.silentProfileRequest, this.profileRequest);
    }

    /**
     * Gets the token to read mail data from MS Graph silently, or falls back to interactive redirect.
     */
    async getMailTokenRedirect() {
        this.silentMailRequest.account = this.account;
        return this.getTokenRedirect(this.silentMailRequest, this.mailRedirectRequest);
    }

    /**
     * Gets the token to read mail data from MS Graph silently, or falls back to interactive popup.
     */
    async getMailTokenPopup() {
        this.silentMailRequest.account = this.account;
        return this.getTokenPopup(this.silentMailRequest, this.mailRequest);
    }

    /**
     * Gets a token silently, or falls back to interactive popup.
     */
    async getTokenPopup(silentRequest, interactiveRequest) {
        try {
            const response = await this.myMSALObj.acquireTokenSilent(silentRequest);
            return response.accessToken;
        } catch (e) {
            console.log("silent token acquisition fails.");
            if (e instanceof InteractionRequiredAuthError) {
                console.log("acquiring token using redirect");
                return this.myMSALObj.acquireTokenPopup(interactiveRequest).then((resp) => {
                    return resp.accessToken;
                }).catch((err) => {
                    console.error(err);
                    return null;
                });
            } else {
                console.error(e);
            }
        }
    }

    /**
     * Gets a token silently, or falls back to interactive redirect.
     */
    async getTokenRedirect(silentRequest, interactiveRequest) {
        try {
            const response = await this.myMSALObj.acquireTokenSilent(silentRequest);
            return response.accessToken;
        } catch (e) {
            console.log("silent token acquisition fails.");
            if (e instanceof InteractionRequiredAuthError) {
                console.log("acquiring token using redirect");
                this.myMSALObj.acquireTokenRedirect(interactiveRequest).catch(console.error);
            } else {
                console.error(e);
            }
        }
    }
}