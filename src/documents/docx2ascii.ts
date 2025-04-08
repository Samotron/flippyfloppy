/**
 * DOCX to AsciiDoc Converter with Built-in Tests
 * 
 * This file contains both the implementation code and the tests in a single file.
 */
import * as fs from 'fs';
import * as path from 'path';
import mammoth from 'mammoth';
import * as turndown from 'turndown';

// ======================================
// IMPLEMENTATION
// ======================================

/**
 * Options for DOCX to AsciiDoc conversion
 */
export interface DocxToAsciidocOptions {
	// Document heading level (0-5)
	headingLevel?: number;
	// Whether to include document attributes in the output
	includeAttributes?: boolean;
	// Custom document attributes to add
	attributes?: Record<string, string>;
	// Whether to keep line breaks
	preserveLineBreaks?: boolean;
	// Style of code blocks ('fenced' or 'indented')
	codeBlockStyle?: 'fenced' | 'indented';
	// Include table of contents
	includeToc?: boolean;
}

/**
 * Converts a DOCX file to AsciiDoc format
 * 
 * @param docxPath - Path to the input DOCX file
 * @param outputPath - Path where the AsciiDoc file should be saved (optional)
 * @param options - Configuration options for the conversion
 * @returns Promise resolving to the AsciiDoc content as a string
 */
export async function convertDocxToAsciidoc(
	docxPath: string,
	outputPath?: string,
	options: DocxToAsciidocOptions = {}
): Promise<string> {
	try {
		// Validate input file exists
		if (!fs.existsSync(docxPath)) {
			throw new Error(`Input DOCX file not found: ${docxPath}`);
		}

		// Step 1: Convert DOCX to HTML using mammoth
		console.log('Converting DOCX to HTML...');
		const { value: html, messages } = await mammoth.convertToHtml({ path: docxPath });

		if (messages.length > 0) {
			console.warn('Conversion warnings:', messages);
		}

		// Step 2: Configure Turndown for HTML to Markdown conversion
		console.log('Converting HTML to Markdown...');
		const turndownService = new turndown({
			headingStyle: 'atx',
			codeBlockStyle: options.codeBlockStyle || 'fenced',
			bulletListMarker: '*',
			emDelimiter: '_',
			strongDelimiter: '*'
		});

		// Handle tables better
		turndownService.addRule('tables', {
			filter: ['table'],
			replacement: function(content, node) {
				// Basic table conversion
				return '\n\n|===\n' + content + '\n|===\n\n';
			}
		});

		// Step 3: Convert HTML to Markdown
		const markdown = turndownService.turndown(html);

		// Step 4: Convert Markdown to AsciiDoc
		console.log('Converting Markdown to AsciiDoc...');
		let asciidoc = await markdownToAsciidoc(markdown, options);

		// Step 5: Post-process AsciiDoc for better formatting
		asciidoc = postProcessAsciidoc(asciidoc, options);

		// Step 6: Write to file if outputPath is provided
		if (outputPath) {
			const outputDir = path.dirname(outputPath);
			if (!fs.existsSync(outputDir)) {
				fs.mkdirSync(outputDir, { recursive: true });
			}

			fs.writeFileSync(outputPath, asciidoc, 'utf8');
			console.log(`Successfully converted DOCX to AsciiDoc: ${outputPath}`);
		}

		return asciidoc;
	} catch (error) {
		console.error('Error converting DOCX to AsciiDoc:', error);
		throw new Error(`Failed to convert DOCX to AsciiDoc: ${error.message}`);
	}
}

/**
 * Converts Markdown content to AsciiDoc
 * 
 * @param markdown - Markdown content to convert
 * @param options - Conversion options
 * @returns AsciiDoc content
 */
async function markdownToAsciidoc(
	markdown: string,
	options: DocxToAsciidocOptions
): Promise<string> {
	let asciidoc = markdown;

	// Convert headers
	asciidoc = asciidoc.replace(/^(#{1,6})\s+(.+)$/gm, (match, hashes, content) => {
		const level = hashes.length + (options.headingLevel || 0);
		return `${'='.repeat(level)} ${content}`;
	});

	// Convert bold
	asciidoc = asciidoc.replace(/\*\*(.*?)\*\*/g, '*$1*');

	// Convert italic
	asciidoc = asciidoc.replace(/_(.*?)_/g, '_$1_');

	// Convert code blocks with language specification
	asciidoc = asciidoc.replace(/```(\w+)\n([\s\S]*?)```/g, (match, lang, code) => {
		return `[source,${lang}]\n----\n${code}----\n`;
	});

	// Convert simple code blocks
	asciidoc = asciidoc.replace(/```\n([\s\S]*?)```/g, '----\n$1----\n');

	// Convert inline code
	asciidoc = asciidoc.replace(/`([^`]+)`/g, '`$1`');

	// Convert block quotes
	asciidoc = asciidoc.replace(/^>\s+(.+)$/gm, '[quote]\n____\n$1\n____\n');

	// Convert links
	asciidoc = asciidoc.replace(/\[(.*?)\]\((.*?)\)/g, 'link:$2[$1]');

	// Convert ordered lists
	asciidoc = asciidoc.replace(/^\d+\.\s+(.+)$/gm, '. $1');

	return asciidoc;
}

/**
 * Post-processes AsciiDoc content for better formatting
 * 
 * @param asciidoc - AsciiDoc content to process
 * @param options - Conversion options
 * @returns Processed AsciiDoc content
 */
function postProcessAsciidoc(
	asciidoc: string,
	options: DocxToAsciidocOptions
): string {
	let result = asciidoc;

	// Add document header if not present
	if (!result.startsWith('=')) {
		result = `= Document Title\n\n${result}`;
	}

	// Add document attributes if requested
	if (options.includeAttributes) {
		let attributes = '';
		attributes += 'Author Name <author@example.com>\n';
		attributes += ':toc: left\n';
		attributes += ':icons: font\n';
		attributes += ':source-highlighter: highlight.js\n';

		// Add custom attributes
		if (options.attributes) {
			for (const [key, value] of Object.entries(options.attributes)) {
				attributes += `:${key}: ${value}\n`;
			}
		}

		// Add table of contents if requested
		if (options.includeToc) {
			attributes += ':toc:\n';
			attributes += ':toclevels: 3\n';
		}

		// Insert attributes after the first heading
		result = result.replace(/^(=+\s+.+)$/m, `$1\n${attributes}`);
	}

	// Clean up line breaks
	if (!options.preserveLineBreaks) {
		// Combine consecutive lines that aren't list items, headers, or other special blocks
		result = result.replace(/([^=\-*.\n])\n([^=\-*.\n\s])/g, '$1 $2');
	}

	// Fix table format issues
	result = result.replace(/\|\n\|/g, '|\n|');

	return result;
}

/**
 * Utility function to validate a DOCX file
 * 
 * @param filePath - Path to the DOCX file
 * @returns Promise resolving to a boolean indicating if file is valid
 */
export async function validateDocxFile(filePath: string): Promise<boolean> {
	try {
		if (!fs.existsSync(filePath)) {
			return false;
		}

		// Try to read the file with mammoth to validate it's a proper DOCX
		await mammoth.extractRawText({ path: filePath });
		return true;
	} catch (error) {
		console.error(`Invalid DOCX file: ${filePath}`, error);
		return false;
	}
}

/**
 * Example usage function
 */
export async function runExample(): Promise<void> {
	try {
		const inputPath = path.join(__dirname, 'input', 'example.docx');
		const outputPath = path.join(__dirname, 'output', 'example.adoc');

		const asciidoc = await convertDocxToAsciidoc(inputPath, outputPath, {
			headingLevel: 0,
			includeAttributes: true,
			includeToc: true,
			preserveLineBreaks: false,
			attributes: {
				'author': 'John Doe',
				'email': 'john@example.com',
				'revdate': new Date().toISOString().split('T')[0]
			}
		});

		console.log('Conversion completed successfully');
		console.log('Preview of generated AsciiDoc:');
		console.log(asciidoc.substring(0, 500) + '...');
	} catch (error) {
		console.error('Example conversion failed:', error);
	}
}

// ======================================
// TESTS
// ======================================

// Only run tests when this file is executed directly with Vitest
if (import.meta.vitest) {
	const { describe, it, expect, beforeEach, afterEach, vi } = import.meta.vitest;

	// Mock dependencies for tests
	vi.mock('fs');
	vi.mock('path');
	vi.mock('mammoth', () => ({
		default: {
			convertToHtml: vi.fn().mockResolvedValue({
				value: '<h1>Test Document</h1><p>This is a test paragraph.</p><ul><li>Item 1</li><li>Item 2</li></ul>',
				messages: []
			}),
			extractRawText: vi.fn().mockResolvedValue({
				value: 'Test Document\nThis is a test paragraph.\n* Item 1\n* Item 2',
				messages: []
			})
		}
	}));

	vi.mock('turndown', () => {
		return {
			default: function() {
				return {
					addRule: vi.fn(),
					turndown: vi.fn().mockReturnValue('# Test Document\n\nThis is a test paragraph.\n\n* Item 1\n* Item 2')
				};
			}
		};
	});

	describe('DOCX to AsciiDoc Converter', () => {
		const testDocxPath = '/test/input/test.docx';
		const testOutputPath = '/test/output/test.adoc';
		const testDir = '/test/output';

		beforeEach(() => {
			// Mock filesystem functions
			vi.mocked(fs.existsSync).mockImplementation((path: fs.PathLike) => {
				return path.toString() === testDocxPath;
			});

			vi.mocked(fs.writeFileSync).mockImplementation(() => undefined);
			vi.mocked(fs.mkdirSync).mockImplementation(() => undefined);

			vi.mocked(path.dirname).mockReturnValue(testDir);
		});

		afterEach(() => {
			vi.clearAllMocks();
		});

		it('should convert DOCX to AsciiDoc', async () => {
			const result = await convertDocxToAsciidoc(testDocxPath);

			// Check that result is a string with expected AsciiDoc content
			expect(result).toBeTypeOf('string');
			expect(result).toContain('= Test Document');
		});

		it('should write to file when outputPath is provided', async () => {
			await convertDocxToAsciidoc(testDocxPath, testOutputPath);

			expect(fs.writeFileSync).toHaveBeenCalledTimes(1);
			expect(fs.writeFileSync).toHaveBeenCalledWith(testOutputPath, expect.any(String), 'utf8');
		});

		it('should create output directory if it does not exist', async () => {
			vi.mocked(fs.existsSync).mockImplementation((path: fs.PathLike) => {
				if (path.toString() === testDocxPath) return true;
				if (path.toString() === testDir) return false;
				return false;
			});

			await convertDocxToAsciidoc(testDocxPath, testOutputPath);

			expect(fs.mkdirSync).toHaveBeenCalledWith(testDir, { recursive: true });
		});

		it('should throw error when input file not found', async () => {
			vi.mocked(fs.existsSync).mockReturnValue(false);

			await expect(convertDocxToAsciidoc('/nonexistent.docx'))
				.rejects.toThrow('Input DOCX file not found');
		});

		it('should validate DOCX files correctly', async () => {
			const validResult = await validateDocxFile(testDocxPath);
			expect(validResult).toBe(true);

			vi.mocked(fs.existsSync).mockReturnValue(false);
			const invalidPathResult = await validateDocxFile('/nonexistent.docx');
			expect(invalidPathResult).toBe(false);

			vi.mocked(fs.existsSync).mockReturnValue(true);
			const mammoth = require('mammoth');
			vi.mocked(mammoth.default.extractRawText).mockRejectedValueOnce(new Error('Invalid file'));
			const invalidFileResult = await validateDocxFile('/invalid.docx');
			expect(invalidFileResult).toBe(false);
		});

		it('should include document attributes when requested', async () => {
			const result = await convertDocxToAsciidoc(testDocxPath, undefined, {
				includeAttributes: true,
				attributes: {
					'author': 'Test Author',
					'email': 'test@example.com'
				}
			});

			expect(result).toContain(':author: Test Author');
			expect(result).toContain(':email: test@example.com');
		});

		it('should include table of contents when requested', async () => {
			const result = await convertDocxToAsciidoc(testDocxPath, undefined, {
				includeAttributes: true,
				includeToc: true
			});

			expect(result).toContain(':toc:');
			expect(result).toContain(':toclevels: 3');
		});

		it('should properly convert various Markdown elements to AsciiDoc', async () => {
			// Override turndown mock for this specific test
			const turndownMock = require('turndown');
			vi.mocked(turndownMock.default).mockImplementationOnce(() => {
				return {
					addRule: vi.fn(),
					turndown: vi.fn().mockReturnValue(`
# Heading 1

## Heading 2

**Bold text** and _italic text_

\`\`\`javascript
function test() {
  return "code";
}
\`\`\`

\`\`\`
Plain code block
\`\`\`

> This is a blockquote

[Link text](https://example.com)

1. Ordered item 1
2. Ordered item 2

* Unordered item 1
* Unordered item 2

\`inline code\`
          `)
				};
			});

			const result = await convertDocxToAsciidoc(testDocxPath);

			// Check various conversions
			expect(result).toContain('= Heading 1');
			expect(result).toContain('== Heading 2');
			expect(result).toContain('*Bold text*');
			expect(result).toContain('_italic text_');
			expect(result).toContain('[source,javascript]');
			expect(result).toContain('----\nPlain code block\n----');
			expect(result).toContain('[quote]');
			expect(result).toContain('link:https://example.com[Link text]');
			expect(result).toContain('. Ordered item 1');
			expect(result).toContain('* Unordered item 1');
			expect(result).toContain('`inline code`');
		});

		it('should handle postProcessing correctly', () => {
			const input = '= Document Title\n\nSome content\nMore content';

			// Test line break handling
			const withLineBreaks = postProcessAsciidoc(input, { preserveLineBreaks: true });
			expect(withLineBreaks).toContain('Some content\nMore content');

			const withoutLineBreaks = postProcessAsciidoc(input, { preserveLineBreaks: false });
			expect(withoutLineBreaks).toContain('Some content More content');

			// Test table formatting
			const tableInput = '|===\n|Cell 1\n|Cell 2\n|===';
			const formattedTable = postProcessAsciidoc(tableInput, {});
			expect(formattedTable).toContain('|Cell 1\n|Cell 2');
		});
	});
}

// Execute the example if this file is run directly with Node.js
if (require.main === module) {
	runExample().catch(console.error);
}
