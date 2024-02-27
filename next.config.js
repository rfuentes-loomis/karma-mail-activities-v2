/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  pageExtensions: ["page.jsx", "page.js"],
  distDir: "build",
  output: "standalone",
  async rewrites() {
    return [
      {
        source: "/outlook",
        destination: "/public/outlook/index.html",
      },
    ];
  },
};

module.exports = nextConfig;
