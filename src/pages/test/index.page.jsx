"use client";
import Box from "@mui/material/Box";
import { useAuthUser } from "../activity/taskpane/index.page";
import { useCallback, useState, useEffect } from "react";
function Home() {
  const [officeIsReady, setOfficeIsReady] = useState(false);
  const [emailItem, setEmailItem] = useState(null);
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
    setOfficeIsReady(true);
  }, [officeIsReady]);

  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);

  return (
    <Box>
      <Box>officeIsReady:{officeIsReady ? "true" : "false"}</Box>
      <Box>emailItem:{emailItem?.itemId}</Box>
      <Box>loadingMsUser:{loadingMsUser? "true" : "false"}</Box>
      <Box>loadingMsUserInitial:{loadingMsUserInitial? "true" : "false"}</Box>

      <Box>currentMsUserIsError:{currentMsUserIsError? "true" : "false"}</Box>

      <Box>
        currentMSUser: <pre>{JSON.stringify(currentMSUser, null, 2)} </pre>
      </Box>
      <Box>
        currentMsUserError:<pre>{JSON.stringify(currentMsUserError, null, 2)} </pre>
      </Box>
    </Box>
  );
}

export default Home;
