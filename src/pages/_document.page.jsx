import { Html, Head, Main, NextScript } from "next/document";
import React, { useState, useCallback, useEffect } from "react";
import Script from "next/script";
export default function Document() {
  return (
    <Html lang="en">
      <Head>
        <meta charSet="UTF-8" />
        <meta httpEquiv="X-UA-Compatible" content="IE=Edge" />
        <Script
          strategy="beforeInteractive"
          type="text/javascript"
          src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
          crossOrigin="anonymous"
        ></Script>
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
