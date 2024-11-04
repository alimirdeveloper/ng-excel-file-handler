import { Component, computed, EventEmitter, input, Output, output } from '@angular/core';
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
      const item = new ExcelColumnClass(a.name, a.idName, a.type, a.options, a.optionsObeservable)
      return item;
    }
    );

  });

  @Output() onDataUploaded = new EventEmitter<any[]>();

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
      const dropdownRange = `${header.name.replace(/\s+/g, '')}_Options!$B$1:$B$${lastRow}`;

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

  private addStyleToColumnHeaders(ws: ExcelJS.Worksheet, headers: string[]) {
    const headerRow = ws.addRow(headers);

    // Style headers
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'FF0000FF' }
      };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      cell.font = { bold: true, size: 16 };
    });
  }

  //#region upload
  onUploadClick() {
    const fileInput = document.querySelector<HTMLInputElement>('#fileInput');
    if (fileInput) {
      fileInput.click(); // Trigger the file input click
    }
  }

  onFileChange(event: Event) {
    debugger
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

  private extractDataFromWorksheet(worksheet: ExcelJS.Worksheet): any[] {
    const headers = this.headersClassified();
    const data: any[] = [];


    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const rowData: any = {};
      headers.forEach((header, index) => {
        const hiddenSheet = worksheet.workbook.getWorksheet(`${header.idName}_Options`);
        let dropdownMapping: {
          [key: string]: string;
        } | undefined;
        if (hiddenSheet)
          dropdownMapping = this.getDropdownMapping(hiddenSheet);
        const cellValue = row.getCell(index + 1).value; // Excel columns are 1-based
        // For dropdown, replace the key with its label
        if (header.type === 'drop-down' && dropdownMapping && dropdownMapping[cellValue?.toString() || '']) {
          rowData[header.name] = dropdownMapping[cellValue?.toString() || '']; // Use the label instead of the key
        } else {
          rowData[header.name] = cellValue; // Regular mapping
        }
      });
      data.push(rowData);
    });
    this.onDataUploaded.emit(data);
    return data;
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


}

export interface ExcelColumn {
  name: string,
  idName: string,
  type: 'text' | 'number' | 'drop-down' | 'Percentage',
  options?: { key: string; value: string }[],
  optionsObeservable?: () => Observable<{ key: string; value: string; }[]>
}

export class ExcelColumnClass {
  constructor(
    public name: string,
    public idName: string,
    public type: 'text' | 'number' | 'drop-down' | 'Percentage',
    public options?: { key: string; value: string }[],
    public optionsObeservable?: () => Observable<{ key: string; value: string; }[]>
  ) {
    this.idName = idName.replace(/\s+/g, '');
  }

}