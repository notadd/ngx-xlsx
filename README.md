# ngx-xlsx

基于[SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx)，Angular导入、导出Excel，支持单个或者多个工作表导出、支持自定义表头、支持自定义工作表名称。

## How to use
### step-1

#### 安装`@notadd/ngx-xlsx`
+ `npm install @notadd/ngx-xlsx -S` 
 
#### 安装依赖
+ `npm install xlsx -S`
+ `npm install file-saver -S`

#### step-2

+ 添加NgxXLSXModule到你的Angular Module

```typescript
  import { NgxXLSXModule } from '@notadd/ngx-xlsx';

  @NgModule({
    imports: [
        ...
        NgxXLSXModule
    ],
    declarations: [AppComponent],
    bootstrap: [AppComponent]
  })
  export class AppModule { }
```

#### step-3

+ 在你的component中引入 `NgxXLSXService`
```typescript
  import { NgxXLSXService } from '@notadd/ngx-xlsx';
```

#### step-4

+ 注入service并在需要的地方调用导出方法 `exportAsExcelFile`
+ 注入service并在需要的地方调用导入方法 `importForExcelFile`


## `exportAsExcelFile` 方法

## 参数

|     参数名    |   类型   | 是否必填 | 默认值 |                                                                   说明                                                                  |
|:-------------|:--------|:--------|:------|:---------------------------------------------------------------------------------------------------------------------------------------|
| json          | any[]    | 必填     |        | 需要导出的数据json    多个工作表导出时数据为二维数组：`[[工作表1],[工作表2],[工作表三]]`    单个工作表导出时数据为一维数组：`[工作表1]` |
| exportOptions | ExportOptions   | 非必填     |    {}    | 导出配置    |    

## Interface

```typescript
export interface ExportOptions {
  /* 文件名， 默认为时间戳 */
  fileName?: string;     

  /* 表头，默认为 json 数组对象的 Object.keys, 长度必须与 json 数组对象的 Object.keys 长度相等 */
  /* 多个工作表可以对应多个表头或单个表头，多个表头为二维数组，与多个工作表匹配 */
  headers?: Array<string | Array<string>>;   

  /* Excel 工作表名称，默认为 'sheet' +索引(从1开始) 多个工作表导出时长度必须与 json 数组长度相等 单个工作表导出时长度必须等于1 */   
  sheetNames?: Array<string>;

  /* 需要合并单元格的数组，['A1:B1'] 和 [['0,0', '0,1']] 等效，两种写法都支持 */
  merges?: Array<string | Array<string>>;
}
```

## `importForExcelFile` 方法

## 参数

|     参数名    |   类型   | 是否必填 | 默认值 |                                                                   说明                                                                  |
|:-------------|:--------|:--------|:------|:---------------------------------------------------------------------------------------------------------------------------------------|
| file          | File    | 必填     |        | 需要导入的 Excel 文件对象 |
| importOptions | ImportOptions   | 非必填     |    {}    | 导入配置    |    

## Interface

```typescript
export interface ImportOptions {
  /* 表头行数，默认为1，当 `headerKeys` 不为 `[]` 时导入的数据会按表头行数截取数组 */
  headerRows?: number;   

  /* Excel 表头对应的导入 `json` 的 `key`，默认为 Excel 表头 */
  /* 多个工作表可以对应多个`key`数组或单个`key`数组，多个`key`数组为二维数组 */
  headerKeys?: Array<string | Array<string>>;
}
```
