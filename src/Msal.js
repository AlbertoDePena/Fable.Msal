import { AuthModule } from "./AuthModule";
import { isIE } from "./Util";

/**
 * @typedef {Object} UserInfo - Graph data about the user.
 * @property {String[]=} businessPhones 
 * @property {String=} displayName
 * @property {String=} givenName
 * @property {String=} id
 * @property {String=} jobTitle
 * @property {String=} mail
 * @property {String=} mobilePhone
 * @property {String=} officeLocation
 * @property {String=} preferredLanguage
 * @property {String=} surname
 * @property {String=} userPrincipalName
 */

/**
 * @typedef {Object} EmailAddress  Graph email address
 * @property {String} address 
 */

/**
 * @typedef {Object} FromEmailAddress - Graph from email address
 * @property {EmailAddress} emailAddress
 */

/**
 * @typedef {Object} MailItem - Mail Item from MS Graph
 * @property {FromEmailAddress} from
 * @property {String} subject
 * @property {String} bodyPreview 
 */

/**
 * @typedef {Object} MailInfo - Mail Info from MS Graph
 * @property {MailItem[]=} value
 */

 /**
  * @typedef {Object} MsalConfig - Msal most common configuration
  * @property {String} clientId
  * @property {String} authority
  * @property {String} redirectUri
  * @property {String} cacheLocation - sessionStorage or localStorage
  * @property {Boolean} storeAuthStateInCookie
  * @property {Boolean} useLoginRedirect - login redirect always used on IE browser
  */

// MS Graph API services
const GRAPH_CONFIG = {
    GRAPH_ME_ENDPT: "https://graph.microsoft.com/v1.0/me",
    GRAPH_MAIL_ENDPT: "https://graph.microsoft.com/v1.0/me/messages"
};

let authModule = undefined;

/**
 * Sign in using login redirect or login popup
 * @param {MsalConfig} config 
 */
export function signIn(config) {
    authModule = new AuthModule(config);

    if (isIE || config.useLoginRedirect) {
        // Load auth module when browser window loads. Only required for redirect flows.
        window.addEventListener("load", async () => {
            authModule.loadAuthModule().then(() => {
                if (authModule.account) { return; }
                authModule.loginRedirect();
            });
        });
    } else {
        authModule.loginPopup();
    }
}

/**
 * Called when user clicks "Sign Out"
 */
export function signOut() {
    authModule.logout();
}

export function getUserName() {
    const account = authModule.getAccount();    
    return account ? account.username : "";
}

/**
 * Get bearer token to call web API
 * @return {Promise<String>} The bearer token
 */
export async function getToken() {
    return isIE ? await authModule.getProfileTokenRedirect() : await authModule.getProfileTokenPopup();
}

/**
 * Get user profile from graph API
 * @return {Promise<UserInfo>}
 */
export async function getProfile() {
    const token = isIE ? await authModule.getProfileTokenRedirect() : await authModule.getProfileTokenPopup();
    if (token && token.length > 0) {
        const graphResponse = await callEndpointWithToken(GRAPH_CONFIG.GRAPH_ME_ENDPT, token);
        return graphResponse;
    }
    return Promise.reject("Failed to acquire profile token");
}

/**
 * Get user mail from graph API
 * @return {Promise<MailInfo>}
 */
export async function getMail() {
    const token = isIE ? await authModule.getMailTokenRedirect() : await authModule.getMailTokenPopup();
    if (token && token.length > 0) {
        const graphResponse = await callEndpointWithToken(GRAPH_CONFIG.GRAPH_MAIL_ENDPT, token);
        return graphResponse;
    }
    return Promise.reject("Failed to acquire mail token");
}

/**
 * Makes an Authorization "Bearer"  request with the given accessToken to the given endpoint.
 * @param {String} endpoint 
 * @param {String} accessToken 
 * @return {Promise<UserInfo | MailInfo>}
 */
async function callEndpointWithToken(endpoint, accessToken) {
    const headers = new Headers();
    
    headers.append("Authorization", `Bearer ${accessToken}`);

    const options = {
        method: "GET",
        headers: headers
    };

    const response = await fetch(endpoint, options);
    return (await response.json());
}