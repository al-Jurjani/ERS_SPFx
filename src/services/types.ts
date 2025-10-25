// defines all the data structures (types) we'll use in our app.
// Think of these as "templates" that describe what our data should look like.

// This represents a single expense submission
export interface IExpenseData {
  // Unique identifier for each expense (we'll generate this)
  id: string;
  
  // Employee information
  employeeName: string;
  employeeEmail: string;
  
  // Expense details
  expenseDate: Date;           // When the expense occurred
  amount: number;              // How much money (as a number, not string)
  category: string;            // Category like "Travel", "Food", etc.
  
  // Receipt information
  receiptFileName: string;     // Name of the uploaded file
  receiptURL: string;          // Where the receipt is stored in SharePoint
  
  // Status tracking
  status: 'Pending' | 'Approved' | 'Rejected';  // Only these 3 values allowed
  submissionDate: Date;        // When the expense was submitted
}

// This represents the result of submitting an expense
export interface ISubmissionResult {
  success: boolean;            // Did it work?
  message: string;             // What happened? (for user feedback)
  expenseId?: string;          // The ID of the created expense (optional)
  error?: string;              // Error details if something went wrong (optional)
}


////
////
////

// Configuration for connecting to Excel and SharePoint
export interface IAppConfiguration {
  siteUrl: string;             // SharePoint site URL
  excelFilePath: string;       // Path to ExpenseData.xlsx
  receiptsFolderPath: string;  // Path to Receipts folder
  excelTableName: string;      // Name of the Excel table (we named it "ExpenseTable") (might be named something else by you)
}

// When we read expenses from Excel, we get them in this format
export interface IExcelRow {
  ID: string;
  EmployeeName: string;
  EmployeeEmail: string;
  ExpenseDate: string;         // Excel stores dates as strings
  Amount: number;
  Category: string;
  ReceiptURL: string;
  Status: string;
}

///
///
///

/**
 * What This Does:
Defines the "shape" of our data
Ensures we don't make typos (TypeScript will warn us)
Makes code easier to understand and maintain
Like creating a blueprint before building a house
 */