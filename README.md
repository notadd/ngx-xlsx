# ngx-xlsx

基于[SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx)，Angular导出Excel，支持单个或者多个工作表导出、支持自定义表头、支持自定义工作表名称。

## How to use
### step-1

#### 安装`@notadd/ngx-xlsx`
+ `npm install @notadd/ngx-xlsx -S` 
 
#### 安装依赖
+ `npm install xlsx -S`
+ `npm install file-saver -S`

#### step-2

+ 添加NgxXLSXModule到你的AppModule

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



## `exportAsExcelFile` 方法

```typescript
    /**
     * export Excel
     * @param {Array<any>} json
     * @param {string} excelFileName
     * @param {Array<string>} headers
     * @param {Array<string>} sheetNames
     */
    public exportAsExcelFile(json: Array<any>, excelFileName: string, headers: Array<string> = null, sheetNames: Array<string> = null): void {
      ...
    }
```
## 参数

|     参数名    |   类型   | 是否必填 | 默认值 |                                                                   说明                                                                  |
|:-------------|:--------|:--------|:------|:---------------------------------------------------------------------------------------------------------------------------------------|
| json          | any[]    | 必填     |        | 需要导出的数据json    多个工作表导出时数据为二维数组：`[[工作表1],[工作表2],[工作表三]]`    单个工作表导出时数据为一维数组：`[工作表1]` |
| excelFileName | string   | 必填     |        | 导出的文件名前缀，后面会追加时间戳                                                                                                      |
| headers       | string[] | 非必填   | []   | 表头，不填时默认为json数组对象的Object.keys,    长度必须与json数组对象的Object.keys长度相等                                              |
| sheetNames    | string[] | 非必填   | []   | Excel工作表名称，不填时默认为'sheet'+索引(从1开始)   多个工作表导出时长度必须与json数组长度相等   单个工作表导出时长度必须等于1         |
| merges        | string[] or string[][] | 非必填   | []  |  需要合并单元格的数组，`['A1:B1']` 和 `[['0,0', '0,1']]` 等效，两种写法都支持      |
