import { Component } from '@angular/core';

import { NgxXLSXService } from '../../projects/notadd/ngx-xlsx/src/public_api';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  sheet: Array<any>;
  sheetHeaders: Array<string>;
  sheetNames: Array<string>;
  sheetMerges: Array<string | Array<string>>;
  sheetMergeHeaders: Array<string>;

  sheets: Array<any>;
  sheetsHeaders: Array<string | Array<string>>;
  sheetsNames: Array<string>;

  constructor(private xlsxService: NgxXLSXService) {
    this.sheets = [];
    for (let j = 0; j < 3; j++) {
      this.sheets[j] = [];
      for (let i = 0; i < 10; i++) {
        this.sheets[j].push({
          Header1: `Row:${i + 1} Cell:1`,
          Header2: `Row:${i + 1} Cell:2`,
          Header3: `Row:${i + 1} Cell:3`,
          Header4: `Row:${i + 1} Cell:4`,
          Header5: `Row:${i + 1} Cell:5`
        });
      }
    }

    this.sheet = [];
    for (let i = 0; i < 10; i++) {
      this.sheet.push({
        Header1: `Row:${i + 1} Cell:1`,
        Header2: `Row:${i + 1} Cell:2`,
        Header3: `Row:${i + 1} Cell:3`,
        Header4: `Row:${i + 1} Cell:4`,
        Header5: `Row:${i + 1} Cell:5`,
        Header6: `Row:${i + 1} Cell:6`,
        Header7: `Row:${i + 1} Cell:7`,
        Header8: `Row:${i + 1} Cell:8`,
        Header9: `Row:${i + 1} Cell:9`,
        Header10: `Row:${i + 1} Cell:10`
      });
    }

    console.log(this.sheet);

    this.sheetHeaders =  ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    this.sheetNames =  ['工作表一'];
    this.sheetMerges = ['A1:B1'];
    /* this is equivalent */
    /* this.sheetMerges = [['0,0', '0,1']]; */

    this.sheetMergeHeaders =  ['一和二', '', '三', '四', '五', '六', '七', '八', '九', '十'];

    this.sheetsHeaders =  [['一', '二', '三', '四', '五'], ['1', '2', '3', '4', '5'], ['one', 'two', 'three', 'four', 'five']];
    // this.sheetsHeaders =  ['one', 'two', 'three', 'four', 'five'];
    this.sheetsNames =  ['工作表一', '工作表二', '工作表三'];
  }

  exportAsXLSXSingle(): void {
    this.xlsxService.exportAsExcelFile(this.sheet, {fileName: 'excel_single', headers: this.sheetHeaders, sheetNames: this.sheetNames});
  }

  exportAsXLSXMultiple(): void {
    this.xlsxService.exportAsExcelFile(this.sheets, {
      fileName: 'excel_multiple',
      headers: this.sheetsHeaders,
      sheetNames: this.sheetsNames
    });
  }

  exportAsXLSXMerge(): void {
    this.xlsxService.exportAsExcelFile(this.sheet, {
      fileName: 'excel_single',
      headers: this.sheetMergeHeaders,
      sheetNames: this.sheetNames,
      merges: this.sheetMerges
    });
  }

  fileExcelUpload(event: any): void {
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const file: File = target.files[0];

    this.xlsxService.importForExcelFile(file)
      .subscribe(console.log);
  }
}
