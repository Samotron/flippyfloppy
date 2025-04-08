// vitest.config.ts
import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		environment: 'node',
		includeSource: ['src/**/*.ts'], // This enables import.meta.vitest access
		coverage: {
			provider: 'v8',
			reporter: ['text', 'html'],
			include: ['src/**/*.ts'],
			exclude: ['**/node_modules/**', '**/dist/**']
		}
	}
});
