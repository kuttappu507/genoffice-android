import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: { target: 'es2020' },
  // dev/preview are only used for local checks and hosted previews; the shipped app is the Capacitor bundle
  server: { host: true, allowedHosts: true },
  preview: { host: true, allowedHosts: true },
});
