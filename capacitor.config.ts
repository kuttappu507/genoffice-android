import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.kuttappu507.genoffice',
  appName: 'GenOffice',
  webDir: 'dist',
  android: {
    allowMixedContent: false,
  },
};

export default config;
