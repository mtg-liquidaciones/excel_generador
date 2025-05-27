// src/utils/fileSystemUtils.js

import fs from 'fs/promises';
import path from 'path';
// import logger from './logger.js'; // Assuming a logger utility

// Placeholder logger - replace with actual logger import
const logger = {
  info: (message) => console.log(`[INFO] fileSystemUtils: ${message}`),
  warn: (message) => console.warn(`[WARN] fileSystemUtils: ${message}`),
  error: (message, error) => console.error(`[ERROR] fileSystemUtils: ${message}`, error || ''),
  debug: (message) => console.log(`[DEBUG] fileSystemUtils: ${message}`),
};

/**
 * Ensures that a directory exists. If it doesn't, it creates it.
 * @param {string} dirPath - The absolute path to the directory.
 * @returns {Promise<void>}
 * @throws {Error} If the directory cannot be created and doesn't exist.
 */
async function ensureDirectoryExists(dirPath) {
  try {
    await fs.mkdir(dirPath, { recursive: true });
    logger.debug(`Directory ensured: ${dirPath}`);
  } catch (error) {
    // In case of an error other than 'EEXIST' (already exists), rethrow.
    // However, { recursive: true } should prevent errors if it already exists.
    if (error.code !== 'EEXIST') {
      logger.error(`Failed to create directory ${dirPath}: ${error.message}`, error);
      throw error; // Rethrow if it's a critical failure
    }
    logger.debug(`Directory already exists or successfully created: ${dirPath}`);
  }
}

/**
 * Safely reads and parses a JSON file.
 * Returns a default value if the file doesn't exist or if JSON parsing fails.
 * @param {string} filePath - The absolute path to the JSON file.
 * @param {object | Array} [defaultValue={}] - The default value to return on failure.
 * @returns {Promise<object|Array>} The parsed JSON data or the default value.
 */
async function safeReadJsonFile(filePath, defaultValue = {}) {
  try {
    const fileContent = await fs.readFile(filePath, 'utf-8');
    const jsonData = JSON.parse(fileContent);
    logger.debug(`Successfully read and parsed JSON from ${filePath}`);
    return jsonData;
  } catch (error) {
    if (error.code === 'ENOENT') {
      logger.warn(`JSON file not found: ${filePath}. Returning default value.`);
    } else if (error instanceof SyntaxError) {
      logger.error(`Error parsing JSON from file ${filePath}: ${error.message}. Returning default value.`);
    } else {
      logger.error(`Error reading file ${filePath}: ${error.message}. Returning default value.`, error);
    }
    return defaultValue;
  }
}

/**
 * Checks if a given path is a directory.
 * @param {string} filePath - The path to check.
 * @returns {Promise<boolean>} True if it's a directory, false otherwise.
 */
async function isDirectory(filePath) {
    try {
        const stats = await fs.stat(filePath);
        return stats.isDirectory();
    } catch (error) {
        if (error.code === 'ENOENT') {
            return false; // Path does not exist
        }
        logger.error(`Error checking if path is directory ${filePath}: ${error.message}`, error);
        return false; // Or rethrow, depending on desired strictness
    }
}

/**
 * Checks if a given path exists.
 * @param {string} filePath - The path to check.
 * @returns {Promise<boolean>} True if it exists, false otherwise.
 */
async function pathExists(filePath) {
    try {
        await fs.access(filePath);
        return true;
    } catch (error) {
        if (error.code === 'ENOENT') {
            return false;
        }
        logger.error(`Error checking path existence ${filePath}: ${error.message}`, error);
        return false; // Or rethrow for other errors
    }
}


export {
  ensureDirectoryExists,
  safeReadJsonFile,
  isDirectory,
  pathExists,
};