const fetch = require('node-fetch').default;

// Azure AD tenant ID and app ID for the service principal
const tenantId = "a4b2de60-9bd7-43fa-8c11-911b09749203";
const clientId = "d308f3c0-4043-4f80-b63f-736feead9fd0"; // This is the App Registration Client ID (Service Principal)

// This function will now dynamically fetch the app role mappings from the service principal
module.exports = async function (context, req) {
    const user = req.body || {};
    const roles = [];
    
    // Get user ID from the access token
    const userId = await getUserIdFromToken(user.accessToken);
    if (!userId) {
        context.res.status(400).json({ error: 'Invalid token or unable to extract user ID' });
        return;
    }

    // Get the available app roles from the service principal dynamically
    const appRoleMappings = await getAppRolesFromServicePrincipal(user.accessToken);
    if (!appRoleMappings) {
        context.res.status(500).json({ error: 'Failed to retrieve app roles' });
        return;
    }

    // Check user's app role assignments against the dynamic mappings
    for (const [roleName, appRoleId] of Object.entries(appRoleMappings)) {
        if (await isUserInAppRole(userId, appRoleId, user.accessToken)) {
            roles.push(roleName);
        }
    }

    context.res.json({
        roles
    });
};

// Function to get user ID from the access token
async function getUserIdFromToken(bearerToken) {
    const url = 'https://graph.microsoft.com/v1.0/me';
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${bearerToken}`
        }
    });

    if (response.status !== 200) {
        return null;
    }

    const graphResponse = await response.json();
    return graphResponse.id;
}

// Function to get app roles dynamically from the service principal
async function getAppRolesFromServicePrincipal(bearerToken) {
    const url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${clientId}'`;
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${bearerToken}`
        },
    });

    if (response.status !== 200) {
        return null;
    }

    const graphResponse = await response.json();
    const servicePrincipal = graphResponse.value[0];

    // Map the app roles to an object { roleName: appRoleId }
    const appRoleMappings = {};
    servicePrincipal.appRoles.forEach(appRole => {
        if (appRole.value) {
            appRoleMappings[appRole.value] = appRole.id;
        }
    });

    return appRoleMappings;
}

// Function to check if the user is assigned a specific app role
async function isUserInAppRole(userId, appRoleId, bearerToken) {
    const url = `https://graph.microsoft.com/v1.0/users/${userId}/appRoleAssignments`;
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${bearerToken}`
        },
    });

    if (response.status !== 200) {
        return false;
    }

    const graphResponse = await response.json();
    const matchingRoles = graphResponse.value.filter(role => role.appRoleId === appRoleId);
    return matchingRoles.length > 0;
}
