"use client";
import React, { useState, useCallback } from "react";
import { ErrorBoundary } from "react-error-boundary";
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
function MyFallbackComponent({ error, resetErrorBoundary }) {
  return (
    <div role="alert">
      <p>Something went wrong:</p>
      <pre>{error.message}</pre>
      <button onClick={resetErrorBoundary}>Try again</button>
    </div>
  );
}
export default function App({ Component, pageProps }) {
  const [loadedOffice, setLoadedOffice] = useState(false);
  return (
    <>
      <ErrorBoundary
        FallbackComponent={MyFallbackComponent}
        onError={(error, errorInfo) => console.log({ error, errorInfo })}
        onReset={() => {
          // reset the state of your app
        }}
      >
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
          type="text/javascript"
          src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
          crossOrigin="anonymous"
          onLoad={() => {
            Office.initialize = (reason) => {
              setLoadedOffice(true);
            };
          }}
        ></Script>
        {/* Local dev hack */}
        {loadedOffice && (
          <Script
            id="assign_replaceState"
            dangerouslySetInnerHTML={{
              __html: "window.history.replaceState = window._replaceState;Office.initialize = function (){}",
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
      </ErrorBoundary>
    </>
  );
}
