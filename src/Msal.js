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

// MS Graph API services
const GRAPH_CONFIG = {
    GRAPH_ME_ENDPT: "https://graph.microsoft.com/v1.0/me",
    GRAPH_MAIL_ENDPT: "https://graph.microsoft.com/v1.0/me/messages"
};

let authModule = undefined;

export function SignInPopup(config) {
    if (isIE) {
        SignInRedirect(config);
    } else {
        authModule = new AuthModule(config);
        authModule.loginPopup();
    }
}

export function SignInRedirect(config) {
    authModule = new AuthModule(config);

    // Load auth module when browser window loads. Only required for redirect flows.
    window.addEventListener("load", async () => {
        authModule.loadAuthModule().then(() => {
           if (authModule.account) { return; }
           authModule.loginRedirect();
        });
    });
}

/**
 * Called when user clicks "Sign Out"
 */
export function SignOut() {
    authModule.logout();
}

export function GetUserName() {
    const account = authModule.getAccount();
    if (!account) {
        return "";
    }
    return account.username;
}

/**
 * Get user profile from graph API
 * @return {Promise<UserInfo>}
 */
export async function GetProfile() {
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
export async function GetMail() {
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
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made at: ' + new Date().toString());

    const response = await fetch(endpoint, options);
    return (await response.json());
}