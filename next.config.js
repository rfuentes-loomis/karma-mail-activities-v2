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
      {
        source: "/api/graphql/:path*",
        destination: `${process.env.DHARMA_PROXY}:path*`,
      },
    ];
  },
};

module.exports = nextConfig;
