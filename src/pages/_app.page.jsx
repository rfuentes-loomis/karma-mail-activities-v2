"use client";
import { useState } from "react";
import Head from "next/head";
import Script from "next/script";
import { QueryClient, QueryClientProvider } from "react-query";
import Theme from "../common/theme";
import { ThemeProvider } from "@mui/material/styles";
import { LocalizationProvider } from "@mui/x-date-pickers";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import { ReactQueryDevtools } from "@tanstack/react-query-devtools";
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
  const [loadedOffice, setLoadedOffice] = useState(false);
  return (
    <>
      {/* https://learn.microsoft.com/en-us/answers/questions/1070090/using-office-javascript-api-in-next-js */}
      {/* Local dev hack */}
      {
        <Script
          id="store_replaceState"
          dangerouslySetInnerHTML={{
            __html: "window._replaceState = window.replaceState",
          }}
        />
      }
      <Script
        async
        type="text/javascript"
        src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
        crossOrigin="anonymous"
        onLoad={() => setLoadedOffice(true)}
      ></Script>
      {/* Local dev hack */}
      {loadedOffice && (
        <Script
          id="assign_replaceState"
          dangerouslySetInnerHTML={{
            __html: "window.history.replaceState = window._replaceState",
          }}
        />
      )}
      <ThemeProvider theme={Theme}>
        <QueryClientProvider client={queryClient}>
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <Loading center isLoading={loadedOffice == false} />
            {loadedOffice ? <Component {...pageProps} /> : null}
          </LocalizationProvider>
        </QueryClientProvider>
      </ThemeProvider>
    </>
  );
}
