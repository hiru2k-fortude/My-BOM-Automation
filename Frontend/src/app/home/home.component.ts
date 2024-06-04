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
    let ConsumptionInput = 0.03;
  
    let verient = this.SelectVerient;
    
  
    let workbook = new ExcelJS.Workbook();
    let worksheet = workbook.addWorksheet('Sheet1');
  
    ExcelData = [
      [
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
        'Consumption (N)',
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
      ],
    ];
  
    for (let i = 1; i < keys.length; i++) {
      this.Template[keys[i]].forEach((element: any) => {
        if (verient === 'PVH-CKNA') {
          wastageInput = 2.5;
        } else if (verient === 'PVH-TUG') {
          wastageInput = 0;
        }
        if ((element['BOM_SECTION']==='Sewing' || element['BOM_SECTION']==='Packing') && verient === 'PVH-TUG'){
          wastageInput = 2;
        }
        ExcelData.push([
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
        ]);
      });
    }
  
    ExcelData.forEach((row, rowIndex) => {
      let highlightRow = false;
      row.forEach((value, colIndex) => {
        let cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
        cell.value = value;

        // set no center cell
        if (![2, 11, 19, 20].includes(colIndex))
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = {
          size: 13
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
  
        if (colIndex === 18 && value === '') {
          highlightRow = true;
        }
  
        if (rowIndex === 0) {
          cell.font = {
            bold: true,
            size: 14,
            color: { argb: 'FFFFFFFF' },
          };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF800000' },
          };
        }
      });
  
      if (highlightRow) {
        row.forEach((_, colIndex) => {
          let cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' },
          };
        });
      }
    });
  
    worksheet.columns.forEach((column:any) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell:any) => {
        let columnLength = cell.text.length;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength + 8;
    });
  
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: ExcelData.length, column: ExcelData[0].length },
    };
  
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
