import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactStrictMode: false,
  async rewrites() {
    return [
      {
        source: '/index.html',
        destination: '/',
      },
    ];
  },
  devIndicators: false,
};

export default nextConfig;
