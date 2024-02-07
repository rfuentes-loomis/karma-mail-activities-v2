/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  pageExtensions: ["page.jsx", "page.js"],
  distDir: "build",
  output: "standalone",
};

module.exports = nextConfig;
