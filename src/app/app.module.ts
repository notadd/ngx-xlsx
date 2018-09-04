import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { NgxXLXSModule } from '../../projects/notadd/ngx-xlsx/src/public_api';

import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    NgxXLXSModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
