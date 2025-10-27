import * as React from 'react';
import styles from './ExpenseForm.module.scss';
import { IExpenseFormProps } from './IExpenseFormProps';
import { 
  TextField, 
  PrimaryButton, 
  Dropdown, 
  IDropdownOption,
  DatePicker,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

// Import our services
import { ExcelService } from '../../../services/ExcelService';
import { FileUploadService } from '../../../services/FileUploadService';
import { IExpenseData } from '../../../services/types';
import { appConfig, featureFlags } from '../../../services/config';

interface IExpenseFormState {
  employeeName: string;
  employeeEmail: string;
  expenseDate: Date | undefined;
  amount: string;
  category: string;
  receiptFile: File | null;
  isSubmitting: boolean;
  submitMessage: string;
  submitMessageType: MessageBarType;
}

export default class ExpenseForm extends React.Component<IExpenseFormProps, IExpenseFormState> {
  
  private fileInputRef = React.createRef<HTMLInputElement>();
  
  // Our service instances
  private excelService: ExcelService;
  private fileUploadService: FileUploadService;

  private categoryOptions: IDropdownOption[] = [
    { key: 'travel', text: 'Travel' },
    { key: 'food', text: 'Food & Dining' },
    { key: 'office', text: 'Office Supplies' },
    { key: 'equipment', text: 'Equipment' },
    { key: 'software', text: 'Software/Subscriptions' },
    { key: 'other', text: 'Other' }
  ];

  constructor(props: IExpenseFormProps) {
    super(props);
    
    this.state = {
      employeeName: '',
      employeeEmail: '',
      expenseDate: undefined,
      amount: '',
      category: '',
      receiptFile: null,
      isSubmitting: false,
      submitMessage: '',
      submitMessageType: MessageBarType.success
    };

    // Initialize services with mock mode from config
    this.excelService = new ExcelService(appConfig, featureFlags.useMockData);
    this.fileUploadService = new FileUploadService(appConfig, featureFlags.useMockData);

    console.log('✨ ExpenseForm component initialized');
    console.log('Mock mode:', featureFlags.useMockData ? 'ON' : 'OFF');
  }

  private handleNameChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ employeeName: newValue || '' });
  };

  private handleEmailChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ employeeEmail: newValue || '' });
  };

  private handleDateChange = (date: Date | null | undefined): void => {
    this.setState({ expenseDate: date || undefined });
  };

  private handleAmountChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    // Only allow numbers and decimal point
    if (newValue === '' || /^\d*\.?\d*$/.test("newValue")) {
      this.setState({ amount: newValue || '' });
    }
  };

  private handleCategoryChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ category: option ? option.key as string : '' });
  };

  private handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0] || null;
    this.setState({ receiptFile: file });
  };

  private validateForm = (): boolean => {
    const { employeeName, employeeEmail, expenseDate, amount, category, receiptFile } = this.state;
    
    if (!employeeName.trim()) {
      this.setState({ 
        submitMessage: 'Please enter your name', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    if (!employeeEmail.trim() || !this.isValidEmail(employeeEmail)) {
      this.setState({ 
        submitMessage: 'Please enter a valid email address', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    if (!expenseDate) {
      this.setState({ 
        submitMessage: 'Please select an expense date', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    if (!amount || parseFloat(amount) <= 0) {
      this.setState({ 
        submitMessage: 'Please enter a valid amount greater than 0', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    if (!category) {
      this.setState({ 
        submitMessage: 'Please select a category', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    if (!receiptFile) {
      this.setState({ 
        submitMessage: 'Please upload a receipt', 
        submitMessageType: MessageBarType.error 
      });
      return false;
    }

    return true;
  };

  private isValidEmail = (email: string): boolean => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  };

  private handleSubmit = async (): Promise<void> => {
    console.log('🚀 Submit button clicked!');
    console.log('==================================================');

    // Clear previous messages
    this.setState({ submitMessage: '' });

    // Validate form
    if (!this.validateForm()) {
      console.log('❌ Form validation failed');
      return;
    }

    console.log('✅ Form validation passed');

    // Set submitting state
    this.setState({ isSubmitting: true });

    try {
      const { employeeName, employeeEmail, expenseDate, amount, category, receiptFile } = this.state;

      // Step 1: Generate unique expense ID
      const expenseId = this.excelService.generateExpenseId();
      console.log('📝 Generated expense ID:', expenseId);

      // Step 2: Validate and upload receipt file
      console.log('📤 Step 1/2: Uploading receipt file...');
      this.fileUploadService.validateFile(receiptFile!);
      const receiptURL = await this.fileUploadService.uploadReceipt(receiptFile!, expenseId);
      console.log('✅ Receipt uploaded successfully!');
      console.log('Receipt URL:', receiptURL);

      // Step 3: Prepare expense data
      const expenseData: IExpenseData = {
        id: expenseId,
        employeeName: employeeName.trim(),
        employeeEmail: employeeEmail.trim(),
        expenseDate: expenseDate!,
        amount: parseFloat(amount),
        category: category,
        receiptFileName: receiptFile!.name,
        receiptURL: receiptURL,
        status: 'Pending',
        submissionDate: new Date()
      };

      console.log('💾 Step 2/2: Saving to Excel...');
      console.log('Expense data to save:', expenseData);

      // Step 4: Add expense to Excel
      const result = await this.excelService.addExpense(expenseData);

      if (result.success) {
        console.log('✅ Expense saved successfully!');
        console.log('Result:', result);

        // Show success message
        this.setState({
          submitMessage: featureFlags.useMockData 
            ? '✅ Expense submitted successfully! (Mock Mode - Check console for details)'
            : '✅ Expense submitted successfully! Finance team will be notified.',
          submitMessageType: MessageBarType.success,
          isSubmitting: false
        });

        // Log summary
        console.log('='.repeat(50));
        console.log('📊 SUBMISSION SUMMARY:');
        console.log('Expense ID:', expenseId);
        console.log('Employee:', employeeName);
        console.log('Amount: $' + amount);
        console.log('Category:', category);
        console.log('Status: Pending');
        console.log('==================================================');

        // Reset form after 3 seconds
        setTimeout(() => {
          this.resetForm();
        }, 3000);

      } else {
        throw new Error(result.message);
      }

    } catch (error) {
      console.error('❌ Error submitting expense:', error);
      this.setState({
        submitMessage: `Error occured`,
        submitMessageType: MessageBarType.error,
        isSubmitting: false
      });
    }
  };

  private resetForm = (): void => {
    console.log('🔄 Resetting form...');
    
    this.setState({
      employeeName: '',
      employeeEmail: '',
      expenseDate: undefined,
      amount: '',
      category: '',
      receiptFile: null,
      submitMessage: ''
    });
    
    // Reset file input
    if (this.fileInputRef.current) {
      this.fileInputRef.current.value = '';
    }

    console.log('✅ Form reset complete');
  };

  public render(): React.ReactElement<IExpenseFormProps> {
    const { 
      employeeName, 
      employeeEmail, 
      expenseDate, 
      amount, 
      category, 
      receiptFile, 
      isSubmitting, 
      submitMessage,
      submitMessageType 
    } = this.state;

    return (
      <div className={styles.expenseForm}>
        <div className={styles.container}>
          <div className={styles.header}>
            <h2>Expense Reimbursement Form</h2>
            <p>Submit your expense for approval</p>
            {featureFlags.useMockData && (
              <div style={{ 
                backgroundColor: '#fff4ce', 
                padding: '10px', 
                borderRadius: '4px', 
                marginTop: '10px',
                fontSize: '13px',
                border: '1px solid #ffb900'
              }}>
                🎭 <strong>Mock Mode Active</strong> - Data will be simulated (not saved to SharePoint)
              </div>
            )}
          </div>

          {submitMessage && (
            <MessageBar 
              messageBarType={submitMessageType}
              onDismiss={() => this.setState({ submitMessage: '' })}
              dismissButtonAriaLabel="Close"
            >
              {submitMessage}
            </MessageBar>
          )}

          <div className={styles.formSection}>
            <TextField
              label="Employee Name"
              required
              placeholder="Enter your full name"
              value={employeeName}
              onChange={this.handleNameChange}
              disabled={isSubmitting}
            />

            <TextField
              label="Email Address"
              required
              placeholder="your.email@company.com"
              value={employeeEmail}
              onChange={this.handleEmailChange}
              disabled={isSubmitting}
              type="email"
            />

            <DatePicker
              label="Expense Date"
              isRequired
              placeholder="Select a date"
              value={expenseDate}
              onSelectDate={this.handleDateChange}
              disabled={isSubmitting}
              maxDate={new Date()}
            />

            <TextField
              label="Amount"
              required
              placeholder="0.00"
              value={amount}
              onChange={this.handleAmountChange}
              disabled={isSubmitting}
              prefix="$"
            />

            <Dropdown
              label="Category"
              required
              placeholder="Select a category"
              options={this.categoryOptions}
              selectedKey={category}
              onChange={this.handleCategoryChange}
              disabled={isSubmitting}
            />

            <div className={styles.fileUpload}>
              <label className={styles.fileLabel}>
                Receipt Upload <span className={styles.required}>*</span>
              </label>
              <input 
                ref={this.fileInputRef}
                type="file"
                accept="image/*,.pdf"
                onChange={this.handleFileChange}
                disabled={isSubmitting}
                className={styles.fileInput}
                aria-label="Upload receipt file"
              />
              {receiptFile && (
                <div className={styles.fileName}>
                  Selected: {receiptFile.name}
                </div>
              )}
            </div>

            <div className={styles.buttonSection}>
              {isSubmitting ? (
                <Spinner 
                  size={SpinnerSize.large} 
                  label="Submitting your expense..." 
                />
              ) : (
                <PrimaryButton
                  text="Submit Expense"
                  onClick={this.handleSubmit}
                  disabled={isSubmitting}
                />
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }
}