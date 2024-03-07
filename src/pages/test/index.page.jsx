"use client";
import Box from "@mui/material/Box";
import { useAuthUser } from "../activity/taskpane/index.page";
import { useCallback, useState, useEffect } from "react";
function Home() {
  const [officeIsReady, setOfficeIsReady] = useState(false);
  const [emailItem, setEmailItem] = useState(null);
  const [userProfile, setUserProfile] = useState(null);
  const [tokenResponse, setTokenResponse] = useState(null);
  const [tokenResponseError, setTokenResponseError] = useState(null);

  const {
    data: currentMSUser,
    isFetching: loadingMsUser,
    isLoading: loadingMsUserInitial,
    isError: currentMsUserIsError,
    error: currentMsUserError,
  } = useAuthUser(officeIsReady);
  const officeOnReadyCallback = useCallback(() => {
    if (officeIsReady) return;
    setEmailItem(Office?.context?.mailbox?.item);
    console.log(Office?.context?.mailbox);
    setUserProfile(Office?.context?.mailbox?.userProfile?.email);
    setOfficeIsReady(true);

    Office.context.mailbox.getUserIdentityTokenAsync((d) => setUserProfile);
  }, [officeIsReady]);

  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);

  const onClick = async () => {
    try {
      const accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false, // ?? changed for outlook mac desktop client from true to false because of error
      });

      setTokenResponse(accessToken);
    } catch (error) {
      setTokenResponseError({ error, msg: "we got an error" });
    }
  };

  return (
    <Box>
      <Box>officeIsReady:{officeIsReady ? "true" : "false"}</Box>
      <Box>emailItem:{emailItem?.itemId}</Box>
      <Box>loadingMsUser:{loadingMsUser ? "true" : "false"}</Box>
      <Box>loadingMsUserInitial:{loadingMsUserInitial ? "true" : "false"}</Box>

      <Box>currentMsUserIsError:{currentMsUserIsError ? "true" : "false"}</Box>

      <Box>
        currentMSUser: <pre>{JSON.stringify(currentMSUser, null, 2)} </pre>
      </Box>
      <Box>
        currentMsUserError:<pre>{JSON.stringify(currentMsUserError, null, 2)} </pre>
      </Box>
      <Box>userProfile:{JSON.stringify(userProfile, null, 2)}</Box>
      <Box>tokenResponseError:{JSON.stringify(tokenResponseError, null, 2)}</Box>
      <Box>tokenResponse:{tokenResponse}</Box>
      <button onClick={onClick}>click</button>
    </Box>
  );
}

export default Home;
