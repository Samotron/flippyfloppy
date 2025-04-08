/**
 * GeoConverter - A Swiss Army knife for geospatial data conversion
 * 
 * This utility provides conversion between various geospatial vector formats using DuckDB:
 * - GeoJSON
 * - Shapefile
 * - KML
 * - GML
 * - TopoJSON
 * - WKT (Well-Known Text)
 * - CSV with coordinates
 * - GeoParquet
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as duckdb from '@duckdb/node-api';

// Type definitions for the supported formats
export type GeoFormat =
	| 'geojson'
	| 'shapefile'
	| 'kml'
	| 'gml'
	| 'topojson'
	| 'wkt'
	| 'csv'
	| 'geoparquet';

// Main converter class
export class GeoConverter {
	private static db: duckdb.Database;
	private static conn: duckdb.Connection;

	/**
	 * Initialize DuckDB connection
	 */
	public static async init(): Promise<void> {
		// Create a new database in memory
		this.db = new duckdb.Database(':memory:');
		this.conn = await this.getConnection();

		// Load spatial extension
		await this.conn.exec('INSTALL spatial; LOAD spatial;');
	}

	/**
	 * Get a DuckDB connection
	 */
	private static getConnection(): Promise<duckdb.Connection> {
		return new Promise((resolve, reject) => {
			this.db.connect((err, conn) => {
				if (err) reject(err);
				else resolve(conn);
			});
		});
	}

	/**
	 * Close DuckDB connection
	 */
	public static async close(): Promise<void> {
		if (this.conn) {
			await this.conn.close();
		}
		if (this.db) {
			await this.db.close();
		}
	}

	/**
	 * Convert geospatial data from one format to another
	 * @param inputPath Path to input file
	 * @param outputPath Path to output file
	 * @param inputFormat Format of input data
	 * @param outputFormat Desired output format
	 */
	public static async convert(
		inputPath: string,
		outputPath: string,
		inputFormat: GeoFormat,
		outputFormat: GeoFormat
	): Promise<void> {
		if (!this.conn) {
			await this.init();
		}

		try {
			// Create a temporary table with the input data
			await this.loadData(inputPath, inputFormat);

			// Export to the output format
			await this.exportData(outputPath, outputFormat);
		} catch (error) {
			throw new Error(`Conversion failed: ${error.message}`);
		}
	}

	/**
	 * Convert from a file to another format with format auto-detection
	 * @param inputFile Path to input file
	 * @param outputFile Path to output file
	 * @param outputFormat Desired output format (if not specified, inferred from extension)
	 */
	public static async convertFile(
		inputFile: string,
		outputFile: string,
		outputFormat?: GeoFormat
	): Promise<void> {
		// Determine input format from file extension
		const inputFormat = this.getFormatFromExtension(path.extname(inputFile));

		// If output format not specified, infer from output file extension
		if (!outputFormat) {
			outputFormat = this.getFormatFromExtension(path.extname(outputFile));
		}

		// Convert
		await this.convert(inputFile, outputFile, inputFormat, outputFormat);
	}

	/**
	 * Get format from file extension
	 * @param ext File extension including dot
	 * @returns Corresponding GeoFormat
	 */
	public static getFormatFromExtension(ext: string): GeoFormat {
		switch (ext.toLowerCase()) {
			case '.json':
			case '.geojson':
				return 'geojson';
			case '.shp':
				return 'shapefile';
			case '.kml':
				return 'kml';
			case '.gml':
				return 'gml';
			case '.topojson':
				return 'topojson';
			case '.wkt':
				return 'wkt';
			case '.csv':
				return 'csv';
			case '.parquet':
			case '.geoparquet':
				return 'geoparquet';
			default:
				throw new Error(`Unsupported file extension: ${ext}`);
		}
	}

	/**
	 * Load data from file into DuckDB table
	 * @param filePath Path to input file
	 * @param format Format of input data
	 */
	private static async loadData(filePath: string, format: GeoFormat): Promise<void> {
		// Drop temporary table if exists
		await this.conn.exec('DROP TABLE IF EXISTS temp_geo_data');

		switch (format) {
			case 'geojson':
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', auto_detect=true);
        `);
				break;
			case 'shapefile':
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='shapefile');
        `);
				break;
			case 'kml':
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='kml');
        `);
				break;
			case 'gml':
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='gml');
        `);
				break;
			case 'csv':
				// Try to auto-detect geometry columns
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='csv', auto_detect=true);
        `);
				break;
			case 'geoparquet':
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='parquet');
        `);
				break;
			case 'wkt':
				// For WKT, we need to read as text and then parse
				await this.conn.exec(`
          CREATE TABLE temp_wkt_raw(wkt_text VARCHAR);
          COPY temp_wkt_raw FROM '${filePath}';
          CREATE TABLE temp_geo_data AS
          SELECT ST_GeomFromWKT(wkt_text) AS geometry FROM temp_wkt_raw;
          DROP TABLE temp_wkt_raw;
        `);
				break;
			case 'topojson':
				// TopoJSON needs special handling
				await this.conn.exec(`
          CREATE TABLE temp_geo_data AS 
          SELECT * FROM ST_READ('${filePath}', format='topojson');
        `);
				break;
			default:
				throw new Error(`Unsupported input format: ${format}`);
		}
	}

	/**
	 * Export data from DuckDB table to file
	 * @param filePath Path to output file
	 * @param format Format for output data
	 */
	private static async exportData(filePath: string, format: GeoFormat): Promise<void> {
		switch (format) {
			case 'geojson':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='GeoJSON'
          );
        `);
				break;
			case 'shapefile':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='ESRI Shapefile'
          );
        `);
				break;
			case 'kml':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='KML'
          );
        `);
				break;
			case 'gml':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='GML'
          );
        `);
				break;
			case 'csv':
				// For CSV, we'll convert geometry to WKT
				await this.conn.exec(`
          CREATE OR REPLACE TABLE temp_geo_csv AS
          SELECT 
            ST_AsText(geometry) AS wkt_geometry,
            *
          EXCLUDE (geometry)
          FROM temp_geo_data;
          
          COPY temp_geo_csv TO '${filePath}' (HEADER, DELIMITER ',');
          
          DROP TABLE temp_geo_csv;
        `);
				break;
			case 'geoparquet':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='Parquet'
          );
        `);
				break;
			case 'wkt':
				await this.conn.exec(`
          COPY (SELECT ST_AsText(geometry) AS wkt FROM temp_geo_data) 
          TO '${filePath}';
        `);
				break;
			case 'topojson':
				await this.conn.exec(`
          CALL ST_WRITE(
            'SELECT * FROM temp_geo_data', 
            '${filePath}', 
            format='TopoJSON'
          );
        `);
				break;
			default:
				throw new Error(`Unsupported output format: ${format}`);
		}
	}

	/**
	 * Utility function to get information about the geometry in a file
	 * @param filePath Path to input file
	 * @param format Format of input data
	 * @returns Promise with geometry information
	 */
	public static async getGeometryInfo(
		filePath: string,
		format: GeoFormat
	): Promise<{ count: number, types: string[], bounds: number[] }> {
		if (!this.conn) {
			await this.init();
		}

		try {
			// Load data
			await this.loadData(filePath, format);

			// Get geometry info
			const countResult = await this.query('SELECT COUNT(*) as count FROM temp_geo_data');
			const typesResult = await this.query(`
        SELECT DISTINCT ST_GeometryType(geometry) as geom_type 
        FROM temp_geo_data
        WHERE geometry IS NOT NULL
      `);
			const boundsResult = await this.query(`
        SELECT 
          MIN(ST_XMin(ST_Envelope(geometry))) as min_x,
          MIN(ST_YMin(ST_Envelope(geometry))) as min_y,
          MAX(ST_XMax(ST_Envelope(geometry))) as max_x,
          MAX(ST_YMax(ST_Envelope(geometry))) as max_y
        FROM temp_geo_data
        WHERE geometry IS NOT NULL
      `);

			return {
				count: countResult[0].count,
				types: typesResult.map(row => row.geom_type),
				bounds: [
					boundsResult[0].min_x,
					boundsResult[0].min_y,
					boundsResult[0].max_x,
					boundsResult[0].max_y
				]
			};
		} catch (error) {
			throw new Error(`Failed to get geometry info: ${error.message}`);
		}
	}

	/**
	 * Run a SQL query on DuckDB
	 * @param sql SQL query
	 * @returns Promise with query results
	 */
	private static async query(sql: string): Promise<any[]> {
		return new Promise((resolve, reject) => {
			this.conn.all(sql, (err, rows) => {
				if (err) reject(err);
				else resolve(rows);
			});
		});
	}
}

// Command-line interface
if (require.main === module) {
	const args = process.argv.slice(2);

	if (args.length < 2 || args.length > 3) {
		console.error('Usage: node geo-converter.js <input-file> <output-file> [output-format]');
		process.exit(1);
	}

	const [inputFile, outputFile, outputFormat] = args;

	(async () => {
		try {
			await GeoConverter.init();

			if (outputFormat) {
				await GeoConverter.convertFile(inputFile, outputFile, outputFormat as GeoFormat);
			} else {
				await GeoConverter.convertFile(inputFile, outputFile);
			}

			console.log(`Converted ${inputFile} to ${outputFile} successfully.`);
			await GeoConverter.close();
		} catch (error) {
			console.error(`Error: ${error.message}`);
			await GeoConverter.close();
			process.exit(1);
		}
	})();
}

// Tests using Vitest
if (import.meta.vitest) {
	const { describe, it, expect, beforeEach, afterEach, vi } = import.meta.vitest;

	describe('GeoConverter', () => {
		const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'geo-converter-test-'));

		// Sample test data
		const testGeoJSON = {
			type: 'FeatureCollection',
			features: [
				{
					type: 'Feature',
					properties: { name: 'Test Point' },
					geometry: {
						type: 'Point',
						coordinates: [0, 0]
					}
				},
				{
					type: 'Feature',
					properties: { name: 'Test Line' },
					geometry: {
						type: 'LineString',
						coordinates: [[0, 0], [1, 1]]
					}
				}
			]
		};

		const testGeoJSONPath = path.join(tempDir, 'test.geojson');

		beforeAll(async () => {
			// Create test data
			fs.writeFileSync(testGeoJSONPath, JSON.stringify(testGeoJSON));

			// Initialize DuckDB
			await GeoConverter.init();
		});

		afterAll(async () => {
			// Close DuckDB connection
			await GeoConverter.close();

			// Clean up test files
			fs.rmSync(tempDir, { recursive: true, force: true });
		});

		it('should get format from file extension', () => {
			expect(GeoConverter.getFormatFromExtension('.geojson')).toBe('geojson');
			expect(GeoConverter.getFormatFromExtension('.shp')).toBe('shapefile');
			expect(GeoConverter.getFormatFromExtension('.parquet')).toBe('geoparquet');
			expect(GeoConverter.getFormatFromExtension('.geoparquet')).toBe('geoparquet');
			expect(() => GeoConverter.getFormatFromExtension('.unknown')).toThrow();
		});

		it('should get geometry info from GeoJSON', async () => {
			const info = await GeoConverter.getGeometryInfo(testGeoJSONPath, 'geojson');

			expect(info.count).toBe(2);
			expect(info.types).toContain('Point');
			expect(info.types).toContain('LineString');
			expect(info.bounds).toEqual([0, 0, 1, 1]);
		});

		it('should convert GeoJSON to WKT', async () => {
			const outputPath = path.join(tempDir, 'output.wkt');
			await GeoConverter.convert(testGeoJSONPath, outputPath, 'geojson', 'wkt');

			expect(fs.existsSync(outputPath)).toBe(true);

			const wktContent = fs.readFileSync(outputPath, 'utf8');
			expect(wktContent).toContain('POINT');
			expect(wktContent).toContain('LINESTRING');
		});

		it('should convert GeoJSON to GeoParquet', async () => {
			const outputPath = path.join(tempDir, 'output.parquet');
			await GeoConverter.convert(testGeoJSONPath, outputPath, 'geojson', 'geoparquet');

			expect(fs.existsSync(outputPath)).toBe(true);

			// Get info from the generated parquet file
			const info = await GeoConverter.getGeometryInfo(outputPath, 'geoparquet');
			expect(info.count).toBe(2);
		});

		it('should convert GeoJSON to CSV', async () => {
			const outputPath = path.join(tempDir, 'output.csv');
			await GeoConverter.convert(testGeoJSONPath, outputPath, 'geojson', 'csv');

			expect(fs.existsSync(outputPath)).toBe(true);

			const csvContent = fs.readFileSync(outputPath, 'utf8');
			expect(csvContent).toContain('wkt_geometry');
			expect(csvContent).toContain('name');
			expect(csvContent).toContain('Test Point');
			expect(csvContent).toContain('POINT');
		});

		it('should handle file conversion with auto-detected formats', async () => {
			const outputPath = path.join(tempDir, 'auto.wkt');
			await GeoConverter.convertFile(testGeoJSONPath, outputPath);

			expect(fs.existsSync(outputPath)).toBe(true);
		});

		it('should throw error for unsupported formats', async () => {
			await expect(GeoConverter.convertFile(
				'nonexistent.xyz',
				'output.abc'
			)).rejects.toThrow();
		});
	});
}

// Export for module usage
