"use client";
import React, { useState, useCallback } from "react";
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
  return (
    <>
      <ThemeProvider theme={Theme}>
        <QueryClientProvider client={queryClient}>
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <Component {...pageProps} />
          </LocalizationProvider>
        </QueryClientProvider>
      </ThemeProvider>
    </>
  );
}
