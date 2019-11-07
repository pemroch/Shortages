import { Component, Inject } from '@angular/core';
import { MAT_DIALOG_DATA } from '@angular/material';

@Component({
    selector: 'template-dialog.component',
    templateUrl: './template-dialog.component.html',
    styleUrls: ['./template-dialog.component.css']
})
export class TemplateDialogComponent {
    data: any;

    constructor(@Inject(MAT_DIALOG_DATA) data: any) { this.data = data; }
}
