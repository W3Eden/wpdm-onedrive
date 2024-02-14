const msalParams = {
    auth: {
        authority: "https://login.microsoftonline.com/<?php echo wpdm_valueof($wpdm_onedrive, 'tenant_id'); ?>/",
        clientId: "<?php echo wpdm_valueof($wpdm_onedrive, 'client_id'); ?>",
        redirectUri: "<?php echo wpdm_valueof($wpdm_onedrive, 'redirect_url'); ?>"
    },
}

const app = new msal.PublicClientApplication(msalParams);

async function getToken(command) {
    let accessToken = "";
    let authParams = null;

    switch (command.type) {
        case "SharePoint":
        case "SharePoint_SelfIssued":
            authParams = { scopes: [`${combine(command.resource, ".default")}`] }; 
            break;
        default:
            break;
    }

    try {
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;
    } catch (e) {
        const resp = await app.loginPopup(authParams);
        app.setActiveAccount(resp.account);

        if (resp.idToken) {
            const resp2 = await app.acquireTokenSilent(authParams);
            accessToken = resp2.accessToken;
        }
    }

    return accessToken;
}