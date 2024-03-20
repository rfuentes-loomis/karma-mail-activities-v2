/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  pageExtensions: ["page.jsx", "page.js"],
  async rewrites() {
    return [
      {
        source: "/outlook",
        destination: "/public/outlook/index.html",
      },
      {
        source: "/api/dharma/:path*",
        destination: `${process.env.DHARMA_PROXY}:path*`,
      },
    ];
  },
};

module.exports = nextConfig;
