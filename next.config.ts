import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  serverExternalPackages: [
    "@prisma/client",
    "bcryptjs",
    "hyperformula",
    "exceljs",
    "pdfkit",
  ],
  outputFileTracingIncludes: {
    "/*": ["./node_modules/generator-function/**/*"],
    "/api/**/*": ["./node_modules/generator-function/**/*"],
    "/spreadsheet/**/*": ["./node_modules/generator-function/**/*"],
  },
};

export default nextConfig;