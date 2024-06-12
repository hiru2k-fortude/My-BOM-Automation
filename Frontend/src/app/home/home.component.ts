import { Component, ElementRef, ViewChild } from '@angular/core';
import { FileUploadService } from '../services/file-upload.service';
import { HttpClient } from '@angular/common/http';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css'],
})
export class HomeComponent {
  // get the Element reference for scroll to top and bottom
  @ViewChild('fileSelect') fileSelect: ElementRef | undefined;
  @ViewChild('buttonElement') buttonElement: ElementRef | undefined;

  constructor(
    private fileUploadService: FileUploadService,
    private http: HttpClient
  ) {}

  Template: any;
  Excelfilename: string = '';

  // Selected file for upload
  selectedFile: File | null = null;
  isButtonDisabled: boolean = false;
  ShowSubmit: boolean = false;
  HideSubmit: boolean = true;
  baseUrl: string =
    localStorage.getItem('UEE-BOM-Automation-backend-baseurl') || '';

  // show info message
  StyleCount: string = '';
  StyleDetails: any = [];
  ShowTable: boolean = false;
  fileVariantInput: boolean = false;
  SelectVerient = '';

  // Reference to file input element
  @ViewChild('fileInput') fileInput: ElementRef<HTMLInputElement> | undefined;

  // Loading indicators
  Loading: boolean = false;
  LoadingDone: boolean = true;
  info: string = '';

  // scroll to to function
  scrollToTop() {
    if (this.fileSelect) {
      const fileSelectElement = this.fileSelect.nativeElement as HTMLElement;
      fileSelectElement.scrollIntoView({ behavior: 'smooth' });
    }
  }

  // scroll to bottom function
  scrollToBottom() {
    if (this.buttonElement) {
      const buttonElement = this.buttonElement.nativeElement as HTMLElement;
      buttonElement.scrollIntoView({ behavior: 'smooth' });
    }
  }

  // Toggle loading indicator
  toggleLoading(): void {
    this.Loading = !this.Loading;
    this.LoadingDone = !this.LoadingDone;
  }

  // Toggle submit button
  toggleSubmit(): void {
    this.ShowSubmit = !this.ShowSubmit;
    this.HideSubmit = !this.HideSubmit;
  }

  // Handle file selection
  onFileSelected(event: Event): void {
    setTimeout(() => this.uploadFile(), 100);

    if (this.ShowSubmit && !this.HideSubmit) this.toggleSubmit();
    this.info = '';
    const inputElement = event.target as HTMLInputElement;
    this.selectedFile = (inputElement.files && inputElement.files[0]) || null;
    this.isButtonDisabled = true;
    this.toggleLoading();
  }

  // Clear file input and selected file
  clearFileInput(): void {
    if (this.fileInput && this.fileInput.nativeElement) {
      this.isButtonDisabled = false;
      this.fileInput.nativeElement.value = ''; // Clear the value of the file input
      this.selectedFile = null; // Clear the selected file
    }
  }

  // Upload file
  async uploadFile(): Promise<void> {
    // reset All
    this.Excelfilename = '';
    this.StyleCount = '';
    this.StyleDetails = [];
    this.ShowTable = false;

    try {
      if (!this.selectedFile) {
        alert('No file selected');
      } else {
        // Upload new file
        this.fileUploadService.uploadFile(this.selectedFile).subscribe(
          (response: any) => {
            this.toggleLoading();
            this.info = '- Request is Successfull -';
            setTimeout(() => {
              this.info = '- Please Fill the DropDown and Submit -';
              setTimeout(() => {
                this.info = '';
                this.Excelfilename = this.selectedFile
                  ? `Selected file : ${this.selectedFile.name}`
                  : 'None';
                this.Template = response.UEETemplate;
                this.showInfo();
                this.fileVariantInput = true;
                this.toggleSubmit();
                this.clearFileInput();
              }, 1000);
            }, 1000);
          },
          (error: any) => {
            this.toggleLoading();
            this.info = '- Error Found... review your Excel File -';
            setTimeout(() => {
              this.info =
                '... Please Double Check the Connection Attempt Once More ...';
              setTimeout(() => {
                this.info = 'Try Again';
                this.clearFileInput();
              }, 3500);
            }, 1500);
          }
        );
      }
    } catch (error: any) {
      alert('-- Invalid Response --');
    }
  }
  downloadFile(): void {
    const keys = Object.keys(this.Template);
    let ExcelData: string[][] = [];
    let wastageInput = 0;
    let ConsumptionInput = 0.003;

    let verient = this.SelectVerient;

    let workbook = new ExcelJS.Workbook();
    let worksheet = workbook.addWorksheet('Sheet1');

    // Header row with merged cells for Consumption (N)
    let headerRow = [
      'Placement ID',
      'Amend (Y/N)',
      'Placement',
      'BOM',
      'Company Season',
      'Division',
      'Department',
      'Cluster',
      'Development Style Number',
      'Garment Way',
      'Colorway Type',
      'Garment Colorways',
      'Product Alternatives',
      'RM Size',
      'Wastage %',
      'Brandix Quote',
      'Comment',
      'Consumption (N)', // Main header for Consumption (N)
      '', // Part 1
      '', // Part 2
      '', // Part 3
      '', // Part 4
      'UOM (YY)',
      'RM Color Code',
      'RM Color Name',
      'Size Scale',
      'Size Split',
      'Main Material (Y/N)',
      'OTT Days',
      'Select Placement to Copy Sizes (Y/N)',
      'Matching Requirement',
      'Repeat Length',
      'Joint Line Requirement',
      'Colorways',
    ];

    // Create input value row based on the specified columns
    let inputValueRow: string[] = headerRow.map((header) => {
      if (
        [
          'Placement',
          'BOM',
          'Company Season',
          'Division',
          'Department',
          'Cluster',
          'Development Style Number',
          'Garment Colorways',
          'Wastage %',
          'Brandix Quote',
          'UOM (YY)',
          'RM Color Code',
          'RM Color Name',
          'Size Scale',
          'Size Split',
          'Main Material (Y/N)',
          'OTT Days',
          'Select Placement to Copy Sizes (Y/N)',
          'Matching Requirement',
          'Repeat Length',
          'Joint Line Requirement',
          'Colorways',
        ].includes(header)
      ) {
        return 'Input value';
      } else {
        return '';
      }
    });

    inputValueRow[17] = 'Input Value (Garment Sizes) - Mandatory';
    worksheet.addRow(inputValueRow);

    // Add mandatory row under the specified columns
    let mandatoryRow: string[] = headerRow.map((header) => {
      if (
        [
          'Placement',
          'BOM',
          'Company Season',
          'Division',
          'Department',
          'Cluster',
          'Development Style Number',
          'Garment Colorways',
          'Brandix Quote',
          'UOM (YY)',
          'RM Color Code',
          'RM Color Name',
        ].includes(header)
      ) {
        return 'Mandatory';
      } else if (header === 'Wastage %') {
        return 'Only enter integers (exclude % symbol)';
      } else {
        return '';
      }
    });
    mandatoryRow[21] = 'All';
    worksheet.addRow(mandatoryRow);

    // Add header row to the worksheet
    worksheet.addRow(headerRow);

    // Merge cells for the "Consumption (N)" header
    worksheet.mergeCells('R1:V1');

    // Merge cells for the "Consumption (N)" header
    worksheet.mergeCells('R3:V3');

    // Style the header row
    worksheet.getRow(3).eachCell((cell, colNumber) => {
      cell.font = { bold: true, size: 14 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' },
      };
    });

    // Add data rows
    for (let i = 1; i < keys.length; i++) {
      this.Template[keys[i]].forEach((element: any) => {
        if (verient === 'PVH-CKNA') {
          wastageInput = 2.5;
        } else if (verient === 'PVH-TUG') {
          wastageInput = 0;
        }
        if (
          (element['BOM_SECTION'] === 'Sewing' ||
            element['BOM_SECTION'] === 'Packing') &&
          verient === 'PVH-TUG'
        ) {
          wastageInput = 2;
        }

        const rowData = [
          '',
          '',
          element['PLACEMENT_NAME'],
          element['Bom'],
          element['SEASON'],
          element['Division'],
          element['Department'],
          'BFF',
          element['STYLE_NO_INDIVDUAL'],
          '',
          '',
          element['GMT_COLOR'],
          '',
          '',
          wastageInput.toString(),
          element['BrandixQuote'],
          '',
          '', // Part 1
          '', // Part 2
          '', // Part 3
          '', // Part 4
          ConsumptionInput.toString(),
          element['uom'],
          element['RM_COLOR_REF'],
          element['RM_COLOR_NAME'],
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          '',
        ];
        const rowIndex = element['IDX'] + 3; // Ensure data starts from the 4th row
        worksheet.getRow(rowIndex).values = rowData; // Directly set the row values
      });
    }

    // Add data to worksheet and style rows
    ExcelData.forEach((row, rowIndex) => {
      let highlightRow = false;
      row.forEach((value, colIndex) => {
        let cell = worksheet.getCell(rowIndex + 4, colIndex + 1);
        cell.value = value;

        if (![2, 11, 19, 20].includes(colIndex)) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        cell.font = { size: 13 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };

        if (colIndex >= 17 && colIndex <= 21 && value === '') {
          highlightRow = true;
        }
      });

      // Highlight specific cells under the "Brandix Quote", "UOM", and "RM Size" headers
      //const headerCellsToHighlight = ['C3', 'D3','E3','F3','G3','H3','I3','L3','R3','S3','T3','U3']; // Cells under headers
      //const mandatoryCellsToHighlight = ['C2', 'D2', 'E2','F2','G2','H2','I2','L2']; // Cell coordinates containing "Mandatory"
      //const inputCellsToHighlight = ['','','',''];

      // Combine the cells to highlight and the mandatory cells
      //  const allCells = [
      //  ...headerCellsToHighlight,
      //  ...mandatoryCellsToHighlight,
      //  ...inputCellsToHighlight,

      // ];

      if (highlightRow) {
        row.forEach((_, colIndex) => {
          let cell = worksheet.getCell(rowIndex + 4, colIndex + 1);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' },
          };
        });
      }
    });

    // Adjust column widths
    worksheet.columns.forEach((column) => {
      if (column.eachCell) {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          let columnLength = cell.text.length;
          if (columnLength > maxLength) {
            maxLength = columnLength;
          }
        });
        column.width = maxLength + 8;
      }
    });

    // Apply auto filter
    worksheet.autoFilter = {
      from: { row: 3, column: 1 },
      to: { row: ExcelData.length + 3, column: headerRow.length },
    };

    // Generate and save the Excel file
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      saveAs(blob, 'PVH Template.xlsx');
    });
  }

  // Submit form
  submitForm(): void {
    if (this.SelectVerient === '') {
      alert('Please select a File Variant');
      return;
    } else {
      this.downloadFile();
    }
  }

  // ChangeFileVariant
  ChangeFileVariant(event: Event): void {
    const target = event.target as HTMLSelectElement;
    this.SelectVerient = target.value;
  }

  showInfo(): void {
    let details = [];
    const keys = Object.keys(this.Template);
    this.StyleCount = `Setting for Styles`;

    for (let i = 1; i < keys.length; i++) {
      details.push({
        style: this.Template[keys[i]][0]['STYLE_NO_INDIVDUAL'],
        season: this.Template[keys[i]][0]['SEASON'],
        bom: this.Template[keys[i]][0]['Bom'],
        devition: this.Template[keys[i]][0]['Division'],
        department: this.Template[keys[i]][0]['Department'],
      });
    }
    this.ShowTable = true;
    this.StyleDetails = details;
  }
}
