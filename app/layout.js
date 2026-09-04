import "./globals.css";

export const metadata = {
  title: "HMO Reporting Dashboard",
  description: "Live Report Dashboard from Google Sheets",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}