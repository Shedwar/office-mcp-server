import {
  ConfidentialClientApplication,
  AuthorizationCodeRequest,
  AuthorizationUrlRequest,
} from "@azure/msal-node";
import http from "http";
import { URL } from "url";
import fs from "fs";
import path from "path";

const TOKEN_CACHE_PATH = path.join(
  process.env.HOME ?? ".",
  ".office-mcp-token-cache.json"
);

const SCOPES = [
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "offline_access",
];

function getMsalClient(): ConfidentialClientApplication {
  const clientId = process.env.AZURE_CLIENT_ID;
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!clientId || !tenantId || !clientSecret) {
    throw new Error(
      "Missing Azure credentials. Set AZURE_CLIENT_ID, AZURE_TENANT_ID, and AZURE_CLIENT_SECRET."
    );
  }

  return new ConfidentialClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientSecret,
    },
  });
}

/** Returns a valid access token, refreshing if necessary */
export async function getAccessToken(): Promise<string> {
  const client = getMsalClient();

  // Try to load cached token
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    const cacheData = fs.readFileSync(TOKEN_CACHE_PATH, "utf-8");
    client.getTokenCache().deserialize(cacheData);
  }

  const accounts = await client.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silentResult = await client.acquireTokenSilent({
        scopes: SCOPES,
        account: accounts[0],
      });
      if (silentResult?.accessToken) {
        // Persist updated cache
        fs.writeFileSync(
          TOKEN_CACHE_PATH,
          client.getTokenCache().serialize()
        );
        return silentResult.accessToken;
      }
    } catch {
      // Fall through to interactive auth
    }
  }

  // Interactive auth via local redirect
  return interactiveLogin(client);
}

async function interactiveLogin(
  client: ConfidentialClientApplication
): Promise<string> {
  const redirectUri = "http://localhost:3456/callback";

  const authCodeUrlParams: AuthorizationUrlRequest = {
    scopes: SCOPES,
    redirectUri,
  };

  const authUrl = await client.getAuthCodeUrl(authCodeUrlParams);

  console.error("\n=== Microsoft 365 Authentication Required ===");
  console.error(`Open this URL in your browser:\n\n${authUrl}\n`);

  // Spin up a temporary HTTP server to catch the redirect
  return new Promise((resolve, reject) => {
    const server = http.createServer(async (req, res) => {
      if (!req.url?.startsWith("/callback")) return;

      const parsed = new URL(req.url, "http://localhost:3456");
      const code = parsed.searchParams.get("code");
      const error = parsed.searchParams.get("error");

      res.writeHead(200, { "Content-Type": "text/html" });
      res.end(
        "<html><body><h2>Authentication complete. You may close this tab.</h2></body></html>"
      );
      server.close();

      if (error || !code) {
        reject(new Error(`Auth error: ${error ?? "no code"}`));
        return;
      }

      try {
        const tokenRequest: AuthorizationCodeRequest = {
          scopes: SCOPES,
          redirectUri,
          code,
        };
        const result = await client.acquireTokenByCode(tokenRequest);
        if (!result?.accessToken) throw new Error("No access token returned");

        // Persist cache
        fs.writeFileSync(
          TOKEN_CACHE_PATH,
          client.getTokenCache().serialize()
        );

        console.error("Authentication successful.\n");
        resolve(result.accessToken);
      } catch (e) {
        reject(e);
      }
    });

    server.listen(3456);
    server.on("error", reject);
  });
}
