export const msalConfig = {
  auth: {
    clientId: "7c9169bf-9829-4b5f-92c6-13466b6fa821", // Replace with your Azure AD app client ID
    authority: "https://login.microsoftonline.com/1b40adf5-38e8-4a55-bc2b-dc507df9913c", // Replace with your tenant ID
    redirectUri: "http://localhost:3000",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Chat.ReadWrite", "Chat.Create"],
};