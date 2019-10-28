import { Injectable } from '@angular/core';
import { Observable, Observer } from 'rxjs';

import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';

export interface Merge {
  s: {
    r: number;
    c: number;
  };

  e: {
    r: number;
    c: number
  };
}

export interface ExportOptions {
  fileName?: string;
  headers?: Array<string | Array<string>>;
  sheetNames?: Array<string>;
  merges?: Array<string | Array<string>>;
}

export interface ImportOptions {
  headerRows?: number;
  headerKeys?: Array<string | Array<string>>;
}

@Injectable()
export class NgxXLSXService {

  private excelType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
  private excelExtension = '.xlsx';

  constructor() { }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const arrayBuffer = new ArrayBuffer(buffer.length);
    const view = new Uint8Array(arrayBuffer);

    for (let i = 0; i !== buffer.length; ++i) {
      view[i] = buffer.charCodeAt(i) & 0xFF;
    }

    const data: Blob = new Blob([arrayBuffer], {
      type: this.excelType
    });

    FileSaver.saveAs(data, fileName + this.excelExtension);
  }

  private numberToChart(i: number): string {
    let chartCode = '';
    i = i + 1;

    while (i > 26) {
      let count = Number.parseInt(`${i / 26}`, 10);
      let remainder = i % 26;
      if (remainder === 0) {
        remainder = 26;
        count --;
        chartCode = String.fromCharCode(64 + Number.parseInt(`${remainder}`, 10)) + chartCode;
      } else {
        chartCode = String.fromCharCode(64 + Number.parseInt(`${remainder}`, 10)) + chartCode;
      }

      i = count;
    }

    chartCode = String.fromCharCode(64 + Number.parseInt(`${i}`, 10)) + chartCode;

    return chartCode;
  }

  private validationHeaders(headers: Array<string>, breakpoint: number): Array<string> {
    const result = [];
    for (let i = 0, length = headers.length; i < length; i += breakpoint) {
      result.push(headers.slice(i, i + breakpoint));
    }
    result.map((_, index) => {
      if (result[index].length !== breakpoint) {
        throw new Error('ngx-xlsx: Parameter "headers" length mismatch');
      }
    });

    return result;
  }

  private builtSheet(data: any, { headers, merges }: { headers?: Array<string>, merges?: Array<string | Array<string>> }): XLSX.WorkSheet {
    const json: Array<any> = JSON.parse(JSON.stringify(data));

    /* add headers rows */
    headers.map(_ => {
      json.unshift({});
    });
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json, {skipHeader: true});

    /* custom header */
    if (headers && headers.length !== 0) {
      if (Array.isArray(headers[0])) {
        for (let i = 0; i < headers.length; i++) {
          for (let j = 0; j < headers[i].length; j++) {
            worksheet[this.numberToChart((j)) + (i + 1)] = { v: headers[i][j] };
          }
        }
      }
    }

    /* custom merge */
    if (merges && merges.length > 0) {
      if (!worksheet['!merges']) {
        worksheet['!merges'] = [];
      }

      merges.map(item => {
        if (Array.isArray(item)) {
          const items: Array<number> = item.toString().split(',') as any as Array<number>;
          const merge: Merge = {
            s: {
              r: items[0],
              c: items[1]
            },
            e: {
              r: items[2],
              c: items[3]
            }
          };

          worksheet['!merges'].push(merge);
        } else {
          worksheet['!merges'].push(XLSX.utils.decode_range(item as any as string));
        }
      });
    }

    return worksheet;
  }

  public exportAsExcelFile(json: Array<any>, {
    fileName = `${new Date().getTime()}`,
    headers = [],
    sheetNames = [],
    merges = []
  }: ExportOptions = {}): void {
    /* slice headers by columns */
    let validHeaders: Array<string> = [];

    /* excelFileName is required */
    if (!fileName) {
      throw new Error('ngx-xlsx: Parameter "fileName" is required');
    }

    /* json is required */
    if (!json || !json.length) {
      throw new Error('ngx-xlsx: Parameter "json" is required');
    }

    /* validate headers length */
    if (headers && headers.length) {
      if (Array.isArray(headers[0])) {
        headers.map((_, index) => {
          const columns = Object.keys(json[index][0]).length;
          validHeaders = this.validationHeaders(headers[index] as Array<string>, columns);
        });
      } else {
        const columns = Object.keys(Array.isArray(json[0]) ? json[0][0] : json[0]).length;
        validHeaders = this.validationHeaders(headers as Array<string>, columns);
      }
    }

    /* validate sheetNames length */
    if (sheetNames && sheetNames.length) {
      if (Array.isArray(json[0]) ? sheetNames.length !== json.length : sheetNames.length !== 1) {
        throw new Error('ngx-xlsx: Parameter "sheetNames" length mismatch');
      }
    }

    /* Workbook Object */
    /* workbook.SheetNames is an ordered list of the sheets in the workbook */
    const workbook: XLSX.WorkBook = { SheetNames: [], Sheets: {} };

    /* multi-sheet */
    if (Array.isArray(json[0])) {
      json.map((data, index) => {
        workbook.SheetNames.push(sheetNames && sheetNames.length ? sheetNames[index] : `Sheet${index + 1}`);

        const sheetHeaders = (Array.isArray(headers[0]) ? headers[index] : headers) as Array<string>;
        validHeaders = this.validationHeaders(sheetHeaders, Object.keys(data[0]).length);
        workbook.Sheets[workbook.SheetNames[index]] = this.builtSheet(data, {headers: validHeaders, merges});
      });
    } else {
      workbook.SheetNames.push(sheetNames && sheetNames.length ? sheetNames[0] : `Sheet${1}`);

      workbook.Sheets[workbook.SheetNames[0]] = this.builtSheet(json, {headers: validHeaders, merges});
    }

    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    this.saveAsExcelFile(excelBuffer, fileName);
  }

  public importForExcelFile(file: File, {headerRows = 1, headerKeys = []}: ImportOptions = {}): Observable<Array<any>> {
    return new Observable<Array<any>>((observer: Observer<Array<any>>) => {
      const result: Array<any> = [];
      const reader: FileReader = new FileReader();
      reader.onload = (event: any) => {
        /* read workbook */
        const bstr: string = event.target.result;
        const workbook: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

        /* grab first sheet */
        workbook.SheetNames.map((sheetName, index) => {
          const worksheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
          const header = (Array.isArray(headerKeys[0]) ? headerKeys[index] : headerKeys) as Array<string>;

          const data: [] = <any>(XLSX.utils.sheet_to_json(worksheet, {
            raw: true,
            header: header.length ? header : void (0)
          }));

          result.push(header.length ? data.slice(headerRows, data.length) : data);
        });

        observer.next(result.length > 1 ? result : result[0]);
        observer.complete();
      };

      reader.onerror = (error: any) => {
        observer.error(error);
      };

      reader.readAsBinaryString(file);
    });
  }
}
