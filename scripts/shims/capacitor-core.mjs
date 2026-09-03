// Minimal stand-ins so src/lib/fileio.ts can be bundled for Node-based tests.
export const Capacitor = { isNativePlatform: () => false, getPlatform: () => 'web' };
export const Filesystem = { writeFile: async () => ({ uri: '' }), readFile: async () => ({ data: '' }) };
export const Directory = { Cache: 'CACHE', Documents: 'DOCUMENTS' };
export const Encoding = { UTF8: 'utf8' };
