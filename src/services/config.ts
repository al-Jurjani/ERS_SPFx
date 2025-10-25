/**
 * Configuration file for the application.
 * Once SharePoint set up, update these values.
 */

import { IAppConfiguration } from './types';

/**
 * Application Configuration
 * 
 * TODO: Update these values once SharePoint site is ready!
 * - siteUrl: The SharePoint site URL
 * - excelFilePath: Path to the ExpenseData.xlsx file
 * - receiptsFolderPath: Path to the Receipts folder
 */
export const appConfig: IAppConfiguration = {
  // Example: 'https://yourschool.sharepoint.com/sites/ExpenseApp'
  siteUrl: 'https://placeholder.sharepoint.com/sites/ExpenseApp',
  
  // Example: '/sites/ExpenseApp/Shared Documents/ExpenseData.xlsx'
  excelFilePath: '/sites/ExpenseApp/Shared Documents/ExpenseData.xlsx',
  
  // Example: '/sites/ExpenseApp/Shared Documents/Receipts'
  receiptsFolderPath: '/sites/ExpenseApp/Shared Documents/Receipts',
  
  // This should match the table name in Excel
  excelTableName: 'ExpenseTable'
};

/**
 * Feature Flags
 * Control which features are enabled
 */
export const featureFlags = {
  // Set to FALSE once SharePoint is ready and you want to test real integration
  useMockData: true,
  
  // Enable detailed console logging for debugging
  enableDebugLogs: true
};

/**
 * HOW TO UPDATE THIS FILE:
 * 
 * 1. Get the SharePoint site URL from your colleague
 * 2. Get the Excel file path (should be in Shared Documents)
 * 3. Get the Receipts folder path
 * 4. Update the values above
 * 5. Set useMockData to FALSE when ready to test with real SharePoint
 */