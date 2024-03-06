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
      {emailItem?.itemId}

      <pre>{JSON.stringify(currentMSUser, null, 2)} </pre>
    </Box>
  );
}

export default Home;
