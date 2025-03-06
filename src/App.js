import React, { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { PublicClientApplication } from "@azure/msal-browser";

const App = () => {
  const { instance, accounts } = useMsal();
  const [chatId, setChatId] = useState(null);

  // Initialize Microsoft Teams SDK
  useEffect(() => {
    microsoftTeams.initialize();
  }, []);

  // Function to handle user login
  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch((error) => {
      console.error("Login failed:", error);
    });
  };

  // Function to create a chat in Microsoft Teams
  const createChat = async () => {
    const tokenRequest = {
      scopes: ["Chat.ReadWrite", "Chat.Create"],
      account: accounts[0],
    };

    try {
      // Acquire an access token
      const response = await instance.acquireTokenSilent(tokenRequest);
      const accessToken = response.accessToken;

      // Define the chat data
      const chatData = {
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('USER_ID')`, // Replace with the user's ID
          },
        ],
      };

      // Call the Microsoft Graph API to create a chat
      const graphResponse = await fetch("https://graph.microsoft.com/v1.0/chats", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(chatData),
      });

      const chat = await graphResponse.json();
      console.log("Chat created:", chat);
      setChatId(chat.id); // Save the chat ID for embedding
    } catch (error) {
      console.error("Error creating chat:", error);
    }
  };

  // Function to embed the chat in an iframe
  const embedChat = () => {
    if (chatId) {
      return (
        <iframe
          src={`https://teams.microsoft.com/l/chat/0/0?chatId=${chatId}`}
          width="100%"
          height="500px"
          frameBorder="0"
          title="Microsoft Teams Chat"
        ></iframe>
      );
    }
    return null;
  };

  return (
    <div>
      <h1>Microsoft Teams Chat Embed</h1>

      {/* Show login button if user is not authenticated */}
      <UnauthenticatedTemplate>
        <button onClick={handleLogin}>Login with Microsoft</button>
      </UnauthenticatedTemplate>

      {/* Show chat creation and embedding options if user is authenticated */}
      <AuthenticatedTemplate>
        <button onClick={createChat}>Create Chat</button>
        {embedChat()}
      </AuthenticatedTemplate>
    </div>
  );
};

export default App;