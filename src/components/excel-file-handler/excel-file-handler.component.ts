import { Component, computed, EventEmitter, input, Output } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { Observable } from 'rxjs';

@Component({
  selector: 'excel-file-handler',
  standalone: true,
  imports: [],
  templateUrl: './excel-file-handler.component.html',
  styleUrl: './excel-file-handler.component.scss'
})
export class ExcelFileHandlerComponent {
 templateName = input<string>('ExcelTemplate');
  sheetName = input<string>('Template');
  headers = input<ExcelColumn[]>([]);
  downloadtip = input<string>('download the template and fill data, please dont change structure of the file');
  uploadtip = input<string>('upload filled template file');
  headersClassified = computed(() => {
    return this.headers().map(a => {
      const item = new ExcelColumnClass(a.name, a.idName, a.type, a.options, a.optionsObeservable, a.mandatory, a.validationPattern,a.unique)
      return item;
    }
    );

  });
  errors: { row: number, column: string, message: string }[] = [];

  @Output() onDataUploaded = new EventEmitter();

  loading = false;
  uploading = false;

  createHeaders = () => {
    return this.headersClassified().map(a => a.name);
  };

  async createExcelTemplate() {

    this.loading = true;

    await this.getOptionsFromObservable();

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet(this.sheetName());

    this.setColumnWidths(ws);

    // Set headers
    const headers = this.createHeaders();

    this.addStyleToColumnHeaders(ws, headers);

    // Apply column configurations based on header types
    this.headersClassified().forEach((header, index) => {
      const columnLetter = String.fromCharCode(65 + index); // Convert index to Excel column letter
      this.setColumnsType(header, workbook, ws, columnLetter);
    });

    // Save the Excel file
    await workbook.xlsx.writeBuffer().then((data) => {
      this.downloadFileToClient(data);
    });

    this.loading = false;

  }



  private setColumnWidths(ws: ExcelJS.Worksheet) {
    this.headersClassified().forEach((header, index) => {
      const column = ws.getColumn(index + 1); // Columns are 1-based in ExcelJS
      switch (header.type) {
        case 'text':
          column.width = 30; // Wider for text columns
          break;
        case 'number':
          column.width = 15;
          break;
        case 'Percentage':
          column.width = 12;
          break;
        case 'drop-down':
          column.width = 20;
          break;
        default:
          column.width = 20;
      }
    });
  }

  async getOptionsFromObservable() {
    // Create an array of Promises that resolves once each observable completes
    const observablesPromises = this.headersClassified().map(header => {
      if (header.optionsObeservable) {
        // Return a Promise that resolves when the observable completes
        return new Promise<void>(resolve => {
          header.optionsObeservable!().subscribe(result => {
            header.options = result;
            resolve(); // Resolve the promise once data is received
          });
        });
      }
      return Promise.resolve(); // For headers without an observable, resolve immediately
    });

    // Wait until all observable Promises are resolved
    await Promise.all(observablesPromises);
  }

  private downloadFileToClient(data: ExcelJS.Buffer) {
    const blob = new Blob([data], { type: 'application/octet-stream' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${this.templateName()}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  }

  private setColumnsType(header: ExcelColumn, workbook: ExcelJS.Workbook, ws: ExcelJS.Worksheet, columnLetter: string) {
    switch (header.type) {
      case 'drop-down':
        // Create a hidden sheet to hold ID-Label mapping
        const hiddenSheet = workbook.addWorksheet(`${header.idName}_Options`);

        // Populate hidden sheet with options
        if (header.options) {
          this.creatDropDownList(header, hiddenSheet, ws, columnLetter);
        }

        // Set hidden sheet visibility after data validation is applied
        hiddenSheet.state = 'visible';
        break;

      case 'number':
        ws.getColumn(columnLetter).numFmt = '0'; // Format as number
        break;

      case 'Percentage':
        ws.getColumn(columnLetter).numFmt = '0.00%'; // Format as percentage
        break;

      case 'text':
      default:
        // No special formatting needed for text type
        break;
    }
  }

  private creatDropDownList(header: ExcelColumn, hiddenSheet: ExcelJS.Worksheet, ws: ExcelJS.Worksheet, columnLetter: string) {
    if (header.options) {
      header.options.forEach((option, optionIndex) => {
        hiddenSheet.getCell(`A${optionIndex + 1}`).value = option.key; // ID
        hiddenSheet.getCell(`B${optionIndex + 1}`).value = option.value; // Label
      });

      // Define the named range for the dropdown options on the hidden sheet
      const lastRow = header.options.length;
      const dropdownRange = `${header.idName.replace(/\s+/g, '')}_Options!$B$1:$B$${lastRow}`;

      // Apply data validation to each cell in the column
      for (let row = 2; row <= 1001; row++) {
        ws.getCell(`${columnLetter}${row}`).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [dropdownRange], // Reference the named range
          showErrorMessage: true,
          errorTitle: 'Invalid Choice',
          error: 'Please select a valid option.'
        };
      }
    }
  }


  //#region upload
  onUploadClick() {
    const fileInput = document.querySelector<HTMLInputElement>('#fileInput');
    this.errors = [];
    if (fileInput) {
      fileInput.click(); // Trigger the file input click
    }
  }

  onFileChange(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      this.readExcelFile(file);
    }
  }
  async readExcelFile(file: File) {
    this.uploading = true; // Set uploading to true while processing the file

    const reader = new FileReader();

    reader.onload = async (e) => {
      const buffer = e.target?.result as ArrayBuffer; // Get the result as ArrayBuffer
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer); // Load the Excel file

      const worksheet = workbook.worksheets[0]; // Access the first worksheet
      this.extractDataFromWorksheet(worksheet);

      this.uploading = false; // Reset uploading state after processing
    };

    reader.onerror = (error) => {
      console.error('Error reading file:', error);
      this.uploading = false; // Reset uploading state on error
    };

    // Read the file as an ArrayBuffer
    reader.readAsArrayBuffer(file);
  }


  private getDropdownMapping(hiddenSheet: ExcelJS.Worksheet): { [key: string]: string } {
    const mapping: { [key: string]: string } = {};

    hiddenSheet.eachRow({ includeEmpty: true }, (row) => {
      const value = row.getCell(1).value as string; // ID column
      const key = row.getCell(2).value as string; // Label column
      if (key && value) {
        mapping[key] = value; // Create a mapping of key to value
      }
    });

    return mapping;
  }

  private extractDataFromWorksheet(worksheet: ExcelJS.Worksheet): any[] {
    const headers = this.headersClassified();
    const data: any[] = [];
    this.errors = []; // Reset errors array

    // Object to track unique values for columns marked as unique
    const uniqueValuesMap: { [columnName: string]: Set<string | number> } = {};

    // Initialize sets for unique columns
    headers.forEach(header => {
      if (header.unique) {
        uniqueValuesMap[header.name] = new Set();
      }
    });

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const rowData: any = {};
      let isValidRow = true;

      headers.forEach((header, index) => {
        const hiddenSheet = worksheet.workbook.getWorksheet(`${header.idName}_Options`);
        let dropdownMapping: { [key: string]: string } | undefined;
        if (hiddenSheet) {
          dropdownMapping = this.getDropdownMapping(hiddenSheet);
        }

        const cell = row.getCell(index + 1);
        const cellValue = cell.value;

        // Check for mandatory fields
        if (header.mandatory && (cellValue === null || cellValue === undefined || cellValue === '')) {
          this.errors.push({
            row: rowNumber,
            column: header.name,
            message: `${header.name} is required`
          });
          isValidRow = false;
          return;
        }

        // Validate against pattern if provided
        if (cellValue && header.validationPattern && typeof cellValue === 'string') {
          if (!header.validationPattern.test(cellValue)) {
            this.errors.push({
              row: rowNumber,
              column: header.name,
              message: `${header.name} has invalid format`
            });
            isValidRow = false;
            return;
          }
        }

        // Check for duplicate values in unique columns
        if (header.unique && cellValue !== null && cellValue !== undefined && cellValue !== '') {
          const stringValue = cellValue.toString();
          if (uniqueValuesMap[header.name].has(stringValue)) {
            this.errors.push({
              row: rowNumber,
              column: header.name,
              message: `${header.name} must be unique (duplicate value found)`
            });
            isValidRow = false;
            return;
          }
          uniqueValuesMap[header.name].add(stringValue);
        }

        // For dropdown, replace the key with its label
        if (header.type === 'drop-down' && dropdownMapping && dropdownMapping[cellValue?.toString() || '']) {
          rowData[header.idName] = dropdownMapping[cellValue?.toString() || ''];
        } else {
          rowData[header.idName] = cellValue;
        }
      });

      if (isValidRow) {
        data.push(rowData);
      }
    });

    this.onDataUploaded.emit({
      data: data,
      errors: this.errors,
      isValid: this.errors.length === 0
    });

    return data;
  }
private addStyleToColumnHeaders(ws: ExcelJS.Worksheet, headers: string[]) {
  const headerRow = ws.addRow(headers);

  // Style headers
  headerRow.eachCell((cell, colNumber) => {
    const header = this.headersClassified()[colNumber - 1]; // Columns are 1-based
    
    // Determine fill color based on properties
    let fillColor = 'FFFFFF00'; // Yellow for normal columns
    
    if (header.mandatory && header.unique) {
      fillColor = 'FF800080'; // Purple for columns that are both mandatory and unique
    } else if (header.mandatory) {
      fillColor = 'FFFF0000'; // Red for mandatory
    } else if (header.unique) {
      fillColor = 'FF0000FF'; // Blue for unique
    }
    
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: fillColor },
      bgColor: { argb: 'FF0000FF' }
    };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    cell.font = { bold: true, size: 16 };
    
    // Add comment explaining column requirements
    let note = '';
    if (header.mandatory) note += 'This field is required.\n';
    if (header.unique) note += 'Values must be unique.';
    if (note) cell.note = note.trim();
  });
}

}
export interface ExcelColumn {
  name: string,
  idName: string,
  type: 'text' | 'number' | 'drop-down' | 'Percentage',
  options?: { key: string; value: string }[],
  optionsObeservable?: () => Observable<{ key: string; value: string; }[]>,
  mandatory?: boolean,
  validationPattern?: RegExp,
  unique?: boolean // Add this new property for duplicate checking
}

export class ExcelColumnClass {
  constructor(
    public name: string,
    public idName: string,
    public type: 'text' | 'number' | 'drop-down' | 'Percentage',
    public options?: { key: string; value: string }[],
    public optionsObeservable?: () => Observable<{ key: string; value: string; }[]>,
    public mandatory: boolean = false,
    public validationPattern?: RegExp,
    public unique: boolean = false // Default to false
  ) {
    this.idName = idName.replace(/\s+/g, '');
  }
}