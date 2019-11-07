import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { MatButtonModule, MatDialogModule } from '@angular/material';

import { AppComponent } from './app.component';
import { TemplateDialogComponent } from './template-dialog/template-dialog.component';

@NgModule({
    declarations: [
        AppComponent,
        TemplateDialogComponent
    ],
    imports: [
        BrowserAnimationsModule,
        BrowserModule,
        FormsModule,
        MatButtonModule,
        MatDialogModule
    ],
    entryComponents: [TemplateDialogComponent],
    bootstrap: [AppComponent]
})
export class AppModule { }
