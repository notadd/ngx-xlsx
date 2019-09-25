import { Injectable } from '@angular/core';

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

    FileSaver.saveAs(data, fileName + '_' + new Date().getTime() + this.excelExtension);
  }

  private numberToChart(i: number): string {
    return String.fromCharCode(65 + i);
  }

  private builtSheet(data: any, { headers, merges }: { headers?: Array<string>, merges?: Array<string | Array<string>> }): XLSX.WorkSheet {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    /* custom header */
    if (headers !== null || headers.length !== 0) {
      for (let i = 0; i < headers.length; i++) {
        worksheet[this.numberToChart((i)) + '1'] = { v: headers[i] };
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

  /**
   * export Excel
   * @param json
   * @param excelFileName
   * @param headers
   * @param sheetNames
   * @param merges
   */
  public exportAsExcelFile(
    json: Array<any>,
    excelFileName: string,
    headers: Array<string> = [],
    sheetNames: Array<string> = [],
    merges: Array<string | Array<string>> = []
  ): void {
    /* excelFileName is required */
    if (!excelFileName) {
      throw new Error('ngx-xlsx: Parameter "excelFileName" is required');
    }

    /* json is required */
    if (!json || !json.length) {
      throw new Error('ngx-xlsx: Parameter "json" is required');
    }

    /* validate headers length */
    if (headers && headers.length) {
      if (headers.length !== Object.keys(Array.isArray(json[0]) ? json[0][0] : json[0]).length) {
        throw new Error('ngx-xlsx: Parameter "headers" length mismatch');
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
        workbook.SheetNames.push(sheetNames ? sheetNames[index] : `sheet${index + 1}`);

        workbook.Sheets[workbook.SheetNames[index]] = this.builtSheet(data, {headers, merges});
      });
    } else {
      workbook.SheetNames.push(sheetNames ? sheetNames[0] : `sheet${0}`);

      workbook.Sheets[workbook.SheetNames[0]] = this.builtSheet(json, {headers, merges});
    }

    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }
}
