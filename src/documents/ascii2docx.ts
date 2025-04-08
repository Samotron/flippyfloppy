/**
 * AsciiDoc to DOCX Converter with Built-in Tests
 * 
 * This file contains both the implementation code and the tests in a single file.
 */
import * as fs from 'fs';
import * as path from 'path';
import Asciidoctor from 'asciidoctor';
import HTMLtoDOCX from 'html-to-docx';

// ======================================
// IMPLEMENTATION
// ======================================

/**
 * Options for DOCX conversion
 */
export interface DocxConversionOptions {
	title?: string;
	margins?: {
		top?: number;
		right?: number;
		bottom?: number;
		left?: number;
	};
	header?: string;
	footer?: string;
	orientation?: 'portrait' | 'landscape';
	author?: string;
	fontSize?: number;
	templatePath?: string;
}

/**
 * Converts AsciiDoc content to DOCX format
 * 
 * @param asciidocContent - The AsciiDoc content to convert
 * @param outputPath - Path where the DOCX file should be saved
 * @param options - Configuration options for the conversion
 * @returns Promise resolving to the path of the created DOCX file
 */
export async function convertAsciidocToDocx(
	asciidocContent: string,
	outputPath: string,
	options: DocxConversionOptions = {}
): Promise<string> {
	try {
		// Initialize Asciidoctor
		const asciidoctor = Asciidoctor();

		// Step 1: Convert AsciiDoc to HTML
		const html = asciidoctor.convert(asciidocContent, {
			standalone: true,
			safe: 'server',
			attributes: {
				showtitle: true,
				icons: 'font',
				'source-highlighter': 'highlight.js'
			}
		});

		// Step 2: Prepare configuration for HTML to DOCX conversion
		const conversionConfig: any = {
			title: options.title || 'Converted Document',
			margins: options.margins || {
				top: 1440,
				right: 1440,
				bottom: 1440,
				left: 1440
			},
			orientation: options.orientation || 'portrait',
		};

		if (options.header) conversionConfig.header = options.header;
		if (options.footer) conversionConfig.footer = options.footer;
		if (options.author) conversionConfig.author = options.author;
		if (options.fontSize) conversionConfig.fontSize = options.fontSize;

		// Step 3: Load template file if provided
		let templateBuffer = null;
		if (options.templatePath && fs.existsSync(options.templatePath)) {
			templateBuffer = fs.readFileSync(options.templatePath);
		}

		// Step 4: Generate DOCX from HTML with optional template
		const docxBuffer = await HTMLtoDOCX(html, templateBuffer, conversionConfig);

		// Step 5: Ensure output directory exists
		const outputDir = path.dirname(outputPath);
		if (!fs.existsSync(outputDir)) {
			fs.mkdirSync(outputDir, { recursive: true });
		}

		// Step 6: Write the DOCX file
		fs.writeFileSync(outputPath, docxBuffer);

		console.log(`Successfully converted AsciiDoc to DOCX: ${outputPath}`);
		return outputPath;
	} catch (error) {
		console.error('Error converting AsciiDoc to DOCX:', error);
		throw new Error(`Failed to convert AsciiDoc to DOCX: ${error.message}`);
	}
}

/**
 * Utility function to read AsciiDoc file content
 */
export function readAsciidocFile(filePath: string): string {
	if (!fs.existsSync(filePath)) {
		throw new Error(`File not found: ${filePath}`);
	}
	return fs.readFileSync(filePath, 'utf8');
}

/**
 * Convert an AsciiDoc file to DOCX
 */
export async function convertFile(
	inputPath: string,
	outputPath: string,
	templatePath?: string
): Promise<string> {
	try {
		const asciidocContent = readAsciidocFile(inputPath);
		return await convertAsciidocToDocx(asciidocContent, outputPath, { templatePath });
	} catch (error) {
		console.error(`Error converting file ${inputPath}:`, error);
		throw error;
	}
}

/**
 * Example usage function
 */
export async function runExample(): Promise<void> {
	const asciidocContent = `
= Document Title
Author Name <author@example.com>
v1.0, 2025-04-08

== Introduction

This is a sample AsciiDoc document.

=== Features

* Simple syntax
* Easy to read
* Converts to many formats

== Code Example

[source,javascript]
----
function hello() {
  console.log("Hello, world!");
}
----
`;

	try {
		const outputPath = path.join(__dirname, 'output', 'example-document.docx');
		await convertAsciidocToDocx(asciidocContent, outputPath);
		console.log('Example conversion completed successfully');
	} catch (error) {
		console.error('Example conversion failed:', error);
	}
}

// ======================================
// TESTS
// ======================================

// Only run tests when this file is executed directly with Vitest
// This pattern allows the tests to be in the same file but not
// interfere with normal usage of the library
if (import.meta.vitest) {
	const { describe, it, expect, beforeEach, afterEach, vi } = import.meta.vitest;

	// Mock dependencies for tests
	vi.mock('fs');
	vi.mock('path');
	vi.mock('asciidoctor', () => {
		return () => ({
			convert: vi.fn().mockReturnValue('<html><body><h1>Test Document</h1></body></html>')
		});
	});
	vi.mock('html-to-docx', () => {
		return vi.fn().mockResolvedValue(Buffer.from('mock docx content'));
	});

	describe('AsciiDoc to DOCX Converter', () => {
		const testDir = '/test/output';
		const templatePath = '/test/templates/template.docx';
		const outputPath = '/test/output/test-output.docx';

		beforeEach(() => {
			// Mock filesystem functions
			vi.mocked(fs.existsSync).mockImplementation((path: string) => {
				if (path === templatePath || path === '/test/sample.adoc') return true;
				return false;
			});

			vi.mocked(fs.readFileSync).mockImplementation((path: any) => {
				if (path === templatePath) return Buffer.from('template content');
				if (path === '/test/sample.adoc') return 'Sample AsciiDoc content';
				return Buffer.from('');
			});

			vi.mocked(fs.writeFileSync).mockImplementation(() => undefined);
			vi.mocked(fs.mkdirSync).mockImplementation(() => undefined);

			vi.mocked(path.dirname).mockReturnValue(testDir);
		});

		afterEach(() => {
			vi.clearAllMocks();
		});

		it('should convert AsciiDoc content to DOCX', async () => {
			const asciidocContent = '= Test Document\n\nThis is a test.';

			const result = await convertAsciidocToDocx(asciidocContent, outputPath);

			expect(result).toBe(outputPath);
			expect(fs.writeFileSync).toHaveBeenCalledTimes(1);
			expect(fs.writeFileSync).toHaveBeenCalledWith(outputPath, expect.any(Buffer));
		});

		it('should use a template file when provided', async () => {
			const asciidocContent = '= Template Test\n\nUsing a template.';
			const options = { templatePath };

			const result = await convertAsciidocToDocx(asciidocContent, outputPath, options);

			expect(result).toBe(outputPath);
			expect(fs.readFileSync).toHaveBeenCalledWith(templatePath);
		});

		it('should handle errors during conversion', async () => {
			// Mock HTML to DOCX to throw an error
			vi.mocked(require('html-to-docx')).mockRejectedValueOnce(new Error('Conversion failed'));

			const asciidocContent = '= Error Test\n\nThis should fail.';

			await expect(convertAsciidocToDocx(asciidocContent, outputPath))
				.rejects.toThrow('Failed to convert AsciiDoc to DOCX: Conversion failed');
		});

		it('should read AsciiDoc file content', () => {
			const content = readAsciidocFile('/test/sample.adoc');
			expect(content).toBe('Sample AsciiDoc content');
		});

		it('should throw error when file not found', () => {
			expect(() => readAsciidocFile('/nonexistent.adoc')).toThrow('File not found');
		});

		it('should convert file from input path to output path', async () => {
			const result = await convertFile('/test/sample.adoc', outputPath);
			expect(result).toBe(outputPath);
		});

		it('should use template when converting a file', async () => {
			const result = await convertFile('/test/sample.adoc', outputPath, templatePath);
			expect(result).toBe(outputPath);
			expect(fs.readFileSync).toHaveBeenCalledWith(templatePath);
		});

		it('should create output directory if it does not exist', async () => {
			const asciidocContent = '= Directory Test\n\nTesting directory creation.';

			await convertAsciidocToDocx(asciidocContent, outputPath);

			expect(fs.mkdirSync).toHaveBeenCalledWith(testDir, { recursive: true });
		});

		it('should apply custom formatting options', async () => {
			const asciidocContent = '= Format Test\n\nTesting custom formatting.';
			const options = {
				title: 'Custom Title',
				author: 'Test Author',
				margins: {
					top: 1000,
					right: 1000,
					bottom: 1000,
					left: 1000
				},
				orientation: 'landscape' as const,
				fontSize: 14,
				header: 'Custom Header',
				footer: 'Page ${pageNumber}'
			};

			const htmlToDOCX = vi.mocked(require('html-to-docx'));

			await convertAsciidocToDocx(asciidocContent, outputPath, options);

			// Verify the correct options were passed to HTMLtoDOCX
			expect(htmlToDOCX).toHaveBeenCalledWith(
				expect.any(String),
				null,
				expect.objectContaining({
					title: 'Custom Title',
					author: 'Test Author',
					margins: options.margins,
					orientation: 'landscape',
					fontSize: 14,
					header: 'Custom Header',
					footer: 'Page ${pageNumber}'
				})
			);
		});
	});
}

// Execute the example if this file is run directly with Node.js
if (require.main === module) {
	runExample().catch(console.error);
}
