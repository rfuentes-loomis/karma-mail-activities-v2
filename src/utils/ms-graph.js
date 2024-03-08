let jwt = require("jsonwebtoken");
let form = require("form-urlencoded");
import * as https from "https";
import { JwksClient } from "jwks-rsa";

const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";
const domain = "graph.microsoft.com";
const version = "v1.0";

function getSigningKeys(header, callback) {
  const client = new JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}

async function verifyToken(token, key, validationOptions) {
  if (!token) return {};
  return new Promise((resolve, reject) => jwt.verify(token, key, validationOptions, (err, decoded) => (err ? reject(() => err) : resolve(decoded))));
}

async function validateJwt(authHeader) {
  if (authHeader) {
    const token = authHeader.split(" ")[1];

    const validationOptions = {
      audience: process.env.AZURE_AD_CLIENTID,
    };

    try {
      const results = await verifyToken(token, getSigningKeys, validationOptions);
      return results;
    } catch (error) {
      throw error;
    }
  }
}
export const getAdminToken = async () => {
  const formParams = {
    client_id: process.env.AZURE_AD_CLIENTID,
    client_secret: process.env.AZURE_AD_SECRET,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  };

  const stsDomain = "https://login.microsoftonline.com";
  const tenant = process.env.AZURE_AD_TENANTID;
  const tokenURLSegment = "oauth2/v2.0/token";
  const encodedForm = form(formParams);

  const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
    method: "POST",
    body: encodedForm,
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
  });
  const json = await tokenResponse.json();
  if (json.error) {
    throw new Error("Error Code: " + json.error + "; Microsoft Graph error " + JSON.stringify(json));
  }

  return json;
};
export async function getGraphData(accessToken, apiUrl, customHeaderProps) {
  return new Promise((resolve, reject) => {
    const options = {
      host: domain,
      path: "/" + version + encodeURI(apiUrl),
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: "Bearer " + accessToken,
        "Cache-Control": "private, no-cache, no-store, must-revalidate",
        Expires: "-1",
        Pragma: "no-cache",
        ...customHeaderProps,
      },
    };

    https
      .get(options, (response) => {
        let body = "";
        response.on("data", (d) => {
          body += d;
        });
        response.on("end", () => {
          // The response from the OData endpoint might be an error, say a
          // 401 if the endpoint requires an access token and it was invalid
          // or expired. But a message is not an error in the call of https.get,
          // so the "on('error', reject)" line below isn't triggered.
          // So, the code distinguishes success (200) messages from error
          // messages and sends a JSON object to the caller with either the
          // requested OData or error information.

          if (response.statusCode === 200) {
            if (apiUrl.indexOf("$value") > 0) {
              resolve(body);
            } else {
              const parsedBody = JSON.parse(body);
              resolve(parsedBody);
            }
          } else {
            // The error body sometimes includes an empty space
            // before the first character, remove it or it causes an error.
            body = body.trim();
            resolve(JSON.parse(body));
          }
        });
      })
      .on("error", reject);
  });
}

export const handleAuth = async () => {
  try {
    // eslint-disable-next-line no-undef
    const accessToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false, // ?? changed for outlook mac desktop client from true to false because of error
    });

    try {
      // eslint-disable-next-line no-undef
      const response = await fetch(process.env.NEXT_PUBLIC_USER_TOKEN_ENDPOINT, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
      });

      const result = await response.json();
      return result;
    } catch (error) {
      // eslint-disable-next-line no-undef
      console.error("Error:", error);
      throw error;
    }
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Error obtaining token", error);
    throw error;
  }

  return null;
};

export async function getAccessToken(authorization) {
  if (!authorization) {
    throw new Error("No Authorization header was found.");
  } else {
    await validateJwt(authorization);
    const scopeName = process.env.AZURE_AD_SCOPES;
    const [, /* schema */ assertion] = authorization.split(" ");

    const tokenScopes = jwt.decode(assertion).scp.split(" ");
    const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
    if (!accessAsUserScope) {
      throw new Error("Missing access_as_user");
    }

    const formParams = {
      client_id: process.env.AZURE_AD_CLIENTID,
      client_secret: process.env.AZURE_AD_SECRET,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: assertion,
      requested_token_use: "on_behalf_of",
      scope: [scopeName].join(" "),
    };

    const stsDomain = "https://login.microsoftonline.com";
    const tenant = "common";
    const tokenURLSegment = "oauth2/v2.0/token";
    const encodedForm = form(formParams);

    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    const json = await tokenResponse.json();
    if (json.error) {
      throw new Error("Error Code: " + json.error + "; Microsoft Graph error " + JSON.stringify(json));
    }
    const graphToken = json.access_token;
    const graphUrlSegment = process.env.GRAPH_URL_SEGMENT || "/me";
    const graphData = await getGraphData(graphToken, graphUrlSegment);
    if (graphData?.code) {
      throw new Error("Error Code: " + graphData?.code + "; Microsoft Graph error " + JSON.stringify(graphData));
    } else {
      return { ...graphData, auth: json };
    }

    return json;
  }
}
