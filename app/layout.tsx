import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Automatización Informes COAP",
  description: "Plataforma de automatización de informes COAP y generación de IA.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es">
      <body>
        <main className="layout-container">
          <header className="header">
            <h1>Automatización de Informes COAP 📊</h1>
            <p style={{ color: 'var(--text-muted)' }}>
              Extracción desde Athena, modelado con plantillas Excel y análisis avanzado de IA por Gemini.
            </p>
          </header>
          {children}
        </main>
      </body>
    </html>
  );
}
