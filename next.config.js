/** @type {import('next').NextConfig} */
const nextConfig = {
  eslint: {
    // Warning ki wajah se build fail nahi hone dega
    ignoreDuringBuilds: true,
  },
};

module.exports = nextConfig;
