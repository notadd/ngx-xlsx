import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { NgxXLSXModule } from '../../projects/notadd/ngx-xlsx/src/public_api';

import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    NgxXLSXModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
