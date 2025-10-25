/**
 * This service handles all interactions with the Excel file:
 * - Adding new expense rows
 * - Reading existing expenses
 * - Updating expense status
 * 
 * Right now it's a MOCK version - it simulates Excel operations.
 * replace this with real Microsoft Graph API calls.
 */

/**
 * five method/functions:
 * 1.) addExpense - has a mock and a real (not implemented yet)
 * 2.) getAllExpenses - has a mock and a real (not yet implemented)
 * 3.) updateExpenseStatus - has a mock and a real (not yet implemented)
 * 4.) generateExpenseId
 * 5.) genereteTestData - private, just for testing
 */

import { IExpenseData, ISubmissionResult, IAppConfiguration, IExcelRow } from './types';

export class ExcelService {
  private config: IAppConfiguration;
  private useMockData: boolean;
  
  // Mock database - simulates Excel rows in memory
  private mockExpenses: IExpenseData[] = [];

  // Constructor
  constructor(config: IAppConfiguration, useMockData: boolean = true) {
    this.config = config;
    this.useMockData = useMockData;

    console.log('📊 ExcelService initialized');
    console.log('Mock mode:', useMockData ? 'ON (simulated Excel)' : 'OFF (real Excel)');
    console.log('Excel file path:', config.excelFilePath);
    console.log('Table name:', config.excelTableName);

    // Initialize with some mock data for testing
    if (useMockData) {
      this.initializeMockData();
    }
  }

  // 1.) method/function to add a new expense
  public async addExpense(expenseData: IExpenseData): Promise<ISubmissionResult> {
    console.log('💾 Adding expense to Excel...');
    console.log('Expense data:', expenseData);

    try {
      if (this.useMockData) {
        return await this.mockAddExpense(expenseData);
      } else {
        return await this.realAddExpense(expenseData);
      }
    } catch (error) {
      console.error('❌ Error adding expense:', error);
      return {
        success: false,
        message: 'Failed to submit expense',
        error: error.message
      };
    }
  }

  // mock version of adding expense - we can remove once real one implemented
  private async mockAddExpense(expenseData: IExpenseData): Promise<ISubmissionResult> {
    console.log('🎭 MOCK: Adding expense to mock database...');

    // Simulate network delay
    await this.simulateDelay(800, 1500);

    // Add to our mock array
    this.mockExpenses.push(expenseData);

    console.log('✅ MOCK: Expense added successfully!');
    console.log('Total expenses in mock DB:', this.mockExpenses.length);
    console.log('All expenses:', this.mockExpenses);

    return {
      success: true,
      message: 'Expense submitted successfully! (Mock mode)',
      expenseId: expenseData.id
    };
  }

  /**
   * REAL VERSION - Actually writes to Excel via Microsoft Graph
   * TODO: Implement this once SharePoint is set up
   */
  private async realAddExpense(expenseData: IExpenseData): Promise<ISubmissionResult> {
    console.log('🚀 REAL: Writing to Excel file...');

    try {
      // TODO: Get access token for Microsoft Graph
      // const accessToken = await this.getAccessToken();

      // TODO: Prepare row data in Excel format
      const rowData = this.prepareExcelRow(expenseData);

      // TODO: Add row to Excel table using Microsoft Graph API
      // const graphUrl = `https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:${this.config.excelFilePath}:/workbook/tables/${this.config.excelTableName}/rows`;
      
      // TODO: Make the API call
      // const response = await fetch(graphUrl, {
      //   method: 'POST',
      //   headers: {
      //     'Authorization': `Bearer ${accessToken}`,
      //     'Content-Type': 'application/json'
      //   },
      //   body: JSON.stringify({
      //     values: [rowData]
      //   })
      // });

      // TODO: Check if successful
      // if (!response.ok) {
      //   throw new Error('Failed to add row to Excel');
      // }

      // For now, throw an error since this isn't implemented
      throw new Error('Real Excel integration not implemented yet. Waiting for SharePoint configuration.');

    } catch (error) {
      console.error('❌ Error writing to Excel:', error);
      throw error;
    }
  }

  // Convert our expense data to Excel row format
  private prepareExcelRow(expense: IExpenseData): any[] {
    return [
      expense.id,
      expense.employeeName,
      expense.employeeEmail,
      expense.expenseDate.toISOString().split('T')[0], // Convert to YYYY-MM-DD
      expense.amount,
      expense.category,
      expense.receiptURL,
      expense.status
    ];
  }

  // 2.) gets all expenses and returns Array of all expenses
  public async getAllExpenses(): Promise<IExpenseData[]> {
    console.log('📖 Reading expenses from Excel...');

    if (this.useMockData) {
      return await this.mockGetAllExpenses();
    } else {
      return await this.realGetAllExpenses();
    }
  }

  // mock version - can remove once properly implemented
  private async mockGetAllExpenses(): Promise<IExpenseData[]> {
    console.log('🎭 MOCK: Fetching expenses from mock database...');

    // Simulate network delay
    await this.simulateDelay(500, 1000);

    console.log('✅ MOCK: Fetched', this.mockExpenses.length, 'expenses');
    return [...this.mockExpenses]; // Return a copy
  }

  /**
   * REAL VERSION - Reads from Excel via Microsoft Graph
   * TODO: Implement this once SharePoint is set up
   */
  private async realGetAllExpenses(): Promise<IExpenseData[]> {
    console.log('🚀 REAL: Reading from Excel file...');

    try {
      // TODO: Get access token
      // const accessToken = await this.getAccessToken();

      // TODO: Read table data from Excel
      // const graphUrl = `https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:${this.config.excelFilePath}:/workbook/tables/${this.config.excelTableName}/rows`;
      
      // TODO: Make the API call
      // const response = await fetch(graphUrl, {
      //   method: 'GET',
      //   headers: {
      //     'Authorization': `Bearer ${accessToken}`
      //   }
      // });

      // TODO: Parse the response
      // const data = await response.json();
      // return this.parseExcelRows(data.value);

      throw new Error('Real Excel reading not implemented yet.');

    } catch (error) {
      console.error('❌ Error reading from Excel:', error);
      throw error;
    }
  }

  // function/method to update an expense status
  public async updateExpenseStatus(expenseId: string, newStatus: 'Approved' | 'Rejected'): Promise<boolean> {
    console.log(`🔄 Updating expense ${expenseId} status to: ${newStatus}`);

    if (this.useMockData) {
      return await this.mockUpdateStatus(expenseId, newStatus);
    } else {
      return await this.realUpdateStatus(expenseId, newStatus);
    }
  }

  // MOCK VERSION - Updates status in mock database
  private async mockUpdateStatus(expenseId: string, newStatus: 'Approved' | 'Rejected'): Promise<boolean> {
    console.log('🎭 MOCK: Updating status in mock database...');

    // Simulate delay
    await this.simulateDelay(500, 1000);

    // Find and update the expense
    const expense = this.mockExpenses.find(e => e.id === expenseId);
    if (expense) {
      expense.status = newStatus;
      console.log('✅ MOCK: Status updated successfully!');
      return true;
    } else {
      console.log('❌ MOCK: Expense not found');
      return false;
    }
  }

  /**
   * REAL VERSION - Updates status in Excel
   * TODO: Implement this once SharePoint is set up
   */
  private async realUpdateStatus(expenseId: string, newStatus: 'Approved' | 'Rejected'): Promise<boolean> {
    console.log('🚀 REAL: Updating status in Excel...');

    try {
      // TODO: Find the row with this expense ID
      // TODO: Update the Status column
      // TODO: Use Microsoft Graph PATCH request

      throw new Error('Real status update not implemented yet.');

    } catch (error) {
      console.error('❌ Error updating status:', error);
      return false;
    }
  }

  // Method/function to generate an expense ID
  public generateExpenseId(): string {
    const timestamp = Date.now();
    const random = Math.floor(Math.random() * 1000);
    return `EXP-${timestamp}-${random}`;
  }

  // private method to Initialize some mock data for testing
  private initializeMockData(): void {
    this.mockExpenses = [
      {
        id: 'EXP-001',
        employeeName: 'John Doe',
        employeeEmail: 'john@example.com',
        expenseDate: new Date('2024-10-20'),
        amount: 150.00,
        category: 'Travel',
        receiptFileName: 'receipt_001.pdf',
        receiptURL: '/sites/test/Receipts/receipt_001.pdf',
        status: 'Pending',
        submissionDate: new Date('2024-10-20')
      },
      {
        id: 'EXP-002',
        employeeName: 'Jane Smith',
        employeeEmail: 'jane@example.com',
        expenseDate: new Date('2024-10-21'),
        amount: 45.50,
        category: 'Food',
        receiptFileName: 'receipt_002.jpg',
        receiptURL: '/sites/test/Receipts/receipt_002.jpg',
        status: 'Approved',
        submissionDate: new Date('2024-10-21')
      }
    ];

    console.log('📝 Initialized with', this.mockExpenses.length, 'mock expenses');
  }

  // Helper function to simulate network delay
  private simulateDelay(minMs: number, maxMs: number): Promise<void> {
    const delay = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
    return new Promise(resolve => setTimeout(resolve, delay));
  }
}

/**
 * HOW TO USE THIS SERVICE:
 * 
 * 1. Create an instance:
 *    const excelService = new ExcelService(config, true);
 * 
 * 2. Add a new expense:
 *    const result = await excelService.addExpense(expenseData);
 * 
 * 3. Get all expenses:
 *    const expenses = await excelService.getAllExpenses();
 * 
 * 4. Update status:
 *    await excelService.updateExpenseStatus('EXP-001', 'Approved');
 */

/**
 * What This Does:

Handles file uploads (currently simulated)
Validates file size and type
Simulates network delay to feel realistic
Has placeholder for real SharePoint upload (we'll fill this later)
Logs everything to console so you can see what's happening
 */