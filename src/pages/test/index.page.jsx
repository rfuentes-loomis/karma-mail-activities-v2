"use client";
import Box from "@mui/material/Box";
import { useCallback, useState, useEffect } from "react";
function Home() {
  const [officeIsReady, setOfficeIsReady] = useState(false);
  const [emailItem, setEmailItem] = useState(null);

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
      {emailItem ? "yes" : "no"}
      <pre>{emailItem}</pre>

      {Office?.context?.mailbox?.item?.itemId}
    </Box>
  );
}

export default Home;
