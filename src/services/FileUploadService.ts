/**
 * handles uploading receipt files to SharePoint.
 * Right now it's a MOCK version - it just simulates the upload.
 * replace this with real Microsoft Graph API calls.
 */

import { IAppConfiguration } from './types';

export class FileUploadService {
  private config: IAppConfiguration;
  private useMockData: boolean;

  // the constructuctor... IAppConfiguration is defined in types.ts
  constructor(config: IAppConfiguration, useMockData: boolean = true) {
    this.config = config;
    this.useMockData = useMockData;

    console.log('📁 FileUploadService initialized');
    console.log('Mock mode:', useMockData ? 'ON (simulated uploads)' : 'OFF (real uploads)');
  }

  // function/method: to upload the reciept to SharePoint
  // returns the url of where the file was uploaded
  public async uploadReceipt(file: File, expenseId: string): Promise<string> {
  console.log('📤 Uploading receipt file...');
  console.log('File name:', file.name);
  console.log('File size:', file.size, 'bytes');
  console.log('Expense ID:', expenseId);

  // MOCK MODE: Just simulate an upload
  if (this.useMockData) {
    return await this.mockUploadReceipt(file, expenseId);
  }

  // REAL MODE: This will be implemented later with Microsoft Graph API
  try {
    return await this.realUploadReceipt(file, expenseId);
  } catch (error) {
    console.error('Real upload failed:', error);
    throw error;
  }
}

  ///
  ///
  ///

  // MOCK VERSION of uploadReciept function above - Simulates uploading a file
  // This is what runs right now since we don't have SharePoint yet
  // IT IS A PRIVATE FUNCTION
  private async mockUploadReceipt(file: File, expenseId: string): Promise<string> {
    console.log('🎭 MOCK: Simulating file upload...');

    // Simulate network delay (1-2 seconds)
    await this.simulateDelay(1000, 2000);

    // Generate a fake URL that looks real
    const fileExtension = file.name.split('.').pop();
    const mockFileName = `receipt_${expenseId}.${fileExtension}`;
    const mockUrl = `${this.config.siteUrl}${this.config.receiptsFolderPath}/${mockFileName}`;

    console.log('✅ MOCK: File "uploaded" successfully!');
    console.log('Mock URL:', mockUrl);

    return mockUrl;
  }

  // REAL VERSION - of uploadReciept function above, but PRIVATE
  // uploadReciept will call this method, once this is implemented
  // TODO: Implement this once SharePoint is set up
  private async realUploadReceipt(file: File, expenseId: string): Promise<string> {
  throw new Error('Real file upload not implemented yet. Waiting for SharePoint configuration.');
}

  ///
  ///
  ///

  // Helper function to simulate network delay
  private simulateDelay(minMs: number, maxMs: number): Promise<void> {
    const delay = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
    return new Promise(resolve => setTimeout(resolve, delay));
  }

  // Validate file before upload
  // param file - The file to validate
  // returns True if valid, throws error if not
  public validateFile(file: File): boolean {
    console.log('🔍 Validating file...');

    // Check if file exists
    if (!file) {
      throw new Error('No file selected');
    }

    // Check file size (max 10MB)
    const maxSize = 10 * 1024 * 1024; // 10MB in bytes
    if (file.size > maxSize) {
      throw new Error('File is too large. Maximum size is 10MB.');
    }

    // Check file type (only images and PDFs)
    const allowedTypes = ['image/jpeg', 'image/jpg', 'image/png', 'image/gif', 'application/pdf'];
if (allowedTypes.indexOf(file.type) === -1) {
      throw new Error('Invalid file type. Please upload an image (JPG, PNG, GIF) or PDF.');
    }

    console.log('✅ File is valid');
    return true;
  }
}

///
///
///

/**
 * HOW TO USE THIS SERVICE:
 * 
 * 1. Create an instance:
 *    const fileService = new FileUploadService(config, true);
 * 
 * 2. Validate the file:
 *    fileService.validateFile(file);
 * 
 * 3. Upload the file:
 *    const url = await fileService.uploadReceipt(file, expenseId);
 * 
 * 4. The URL can now be saved to Excel!
 */

///
///
///

/**
What This Does:

Handles file uploads (currently simulated)
Validates file size and type
Simulates network delay to feel realistic
Has placeholder for real SharePoint upload (we'll fill this later)
Logs everything to console so you can see what's happening
 */