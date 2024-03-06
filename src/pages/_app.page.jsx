"use client";
import React, { useState, useCallback, useEffect } from "react";
import Head from "next/head";
import Script from "next/script";
import { QueryClient, QueryClientProvider } from "react-query";
import Theme from "../common/theme";
import { ThemeProvider } from "@mui/material/styles";
import { LocalizationProvider } from "@mui/x-date-pickers";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import { ReactQueryDevtools } from "@tanstack/react-query-devtools/production";
import Loading from "../common/loading";

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      retry: false,
    },
    mutations: {
      retry: false,
    },
  },
});

export default function App({ Component, pageProps }) {
  const [officeIsReady, setOfficeIsReady] = useState(false);

  //#region Handle Office On Ready & set Default values
  const officeOnReadyCallback = useCallback(() => {
    if (officeIsReady) return;
    setOfficeIsReady(true);
  }, [officeIsReady]);
  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);

  //#endregion
  return (
    <>

      <ThemeProvider theme={Theme}>
        <QueryClientProvider client={queryClient}>
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <Loading center isLoading={officeIsReady == false} />
            {officeIsReady ? <Component {...pageProps} /> : <></>}
          </LocalizationProvider>
        </QueryClientProvider>
      </ThemeProvider>
    </>
  );
}
