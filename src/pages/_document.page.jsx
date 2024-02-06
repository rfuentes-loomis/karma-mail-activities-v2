import { Html, Head, Main, NextScript } from "next/document";

export default function Document() {
  return (
    <Html lang="en">
      <Head>
        {/* https://learn.microsoft.com/en-us/answers/questions/1070090/using-office-javascript-api-in-next-js */}
        <meta charSet="UTF-8" />
        <meta httpEquiv="X-UA-Compatible" content="IE=Edge" />
        {process.env.NODE_ENV !== "production" && (
          <script
            dangerouslySetInnerHTML={{
              __html: `window._replaceState = window.replaceState`,
            }}
          />
        )}
        <script
          type="text/javascript"
          src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
        ></script>
        {process.env.NODE_ENV !== "production" && (
          <script
            dangerouslySetInnerHTML={{
              __html: `window.history.replaceState = window._replaceState`,
            }}
          />
        )}
      </Head>
      <body>
        <Main />
        <NextScript />
      </body>
    </Html>
  );
}
