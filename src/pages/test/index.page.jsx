"use client";
import Box from "@mui/material/Box";
import { useAuthUser, useWorkEffortTypes } from "../activity/taskpane/index.page";
import { useCallback, useState, useEffect } from "react";
function Home() {
  const [officeIsReady, setOfficeIsReady] = useState(false);
  const [emailItem, setEmailItem] = useState(null);
  const [userProfile, setUserProfile] = useState(null);
  const [tokenResponse, setTokenResponse] = useState(null);
  const [tokenResponseError, setTokenResponseError] = useState(null);
  const { data: workEffortTypes, isLoading: workEffortTypesLoading } = useWorkEffortTypes();
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
    setUserProfile(Office?.context?.mailbox?.userProfile?.emailAddress);
    setOfficeIsReady(true);

    Office.context.mailbox.getUserIdentityTokenAsync((d) => setUserProfile);
  }, [officeIsReady]);

  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);

  const onClick = async () => {
    try {
      setTokenResponse(Office.mailbox);
    } catch (error) {
      setTokenResponseError({ error, msg: "we got an error" });
    }
  };

  return (
    <Box>
      <Box>officeIsReady:{officeIsReady ? "true" : "false"}</Box>
      <Box>
        emailItem:<pre>{JSON.stringify(emailItem, null, 2)} </pre>
      </Box>
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
      <Box>userProfile:{userProfile}</Box>
      <Box>tokenResponseError:{JSON.stringify(tokenResponseError, null, 2)}</Box>
      <Box>tokenResponse:{tokenResponse}</Box>
      <button onClick={onClick}>click</button>
    </Box>
  );
}

export default Home;
