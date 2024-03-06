import { Html, Head, Main, NextScript } from "next/document";
import React, { useState, useCallback, useEffect } from "react";
import Script from "next/script";
export default function Document() {
  return (
    <Html lang="en">
      <Head>
        <meta charSet="UTF-8" />
        <meta httpEquiv="X-UA-Compatible" content="IE=Edge" />
        {/* https://learn.microsoft.com/en-us/answers/questions/1070090/using-office-javascript-api-in-next-js */}
        {/* Local dev hack */}
        {process.env.NODE_ENV !== "production" && (
          <Script
            id="store_replaceState"
            dangerouslySetInnerHTML={{
              __html: "window._replaceState = window.replaceState",
            }}
          />
        )}
        <Script
          strategy="beforeInteractive"
          defer={false}
          type="text/javascript"
          src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
          crossOrigin="anonymous"
        ></Script>

        {/* Local dev hack */}
        {process.env.NODE_ENV !== "production" && (
          <Script
            id="assign_replaceState"
            dangerouslySetInnerHTML={{
              __html: "window.history.replaceState = window._replaceState",
            }}
          />
        )}

        <Script
          id="office_init"
          dangerouslySetInnerHTML={{
            __html: "Office.initialize = function (){}",
          }}
        />
      </Head>
      <body>
        <Main />
        <NextScript />
      </body>
    </Html>
  );
}
