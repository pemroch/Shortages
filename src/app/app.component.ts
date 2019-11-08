import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material';
import * as XLSX from 'xlsx';

import { TemplateDialogComponent } from './template-dialog/template-dialog.component';

type AOA = any[][];

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
    lowesLinkDailyRoutesFileName: string = '';
    plantPartnerFileName: string = '';

    lowesLinkData: any = [];
    dailyRoutesData: any = [];
    plantPartnerData: any = [];

    wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };

    constructor(private matDialog: MatDialog) { }

    ngOnInit() {
        const lowesLinkDailyRoutesFile = document.getElementById('lowes-link-daily-routes-file')
        lowesLinkDailyRoutesFile.addEventListener("dragover", this.fileDrag, false);
        lowesLinkDailyRoutesFile.addEventListener("dragleave", this.fileDrag, false);
        lowesLinkDailyRoutesFile.addEventListener("drop", (e) => this.lowesLinkDailyRoutesFileDrop(e), false);

        const plantPartnerFile = document.getElementById('plant-partner-file')
        plantPartnerFile.addEventListener("dragover", this.fileDrag, false);
        plantPartnerFile.addEventListener("dragleave", this.fileDrag, false);
        plantPartnerFile.addEventListener("drop", (e) => this.plantPartnerFileDrop(e), false);

        const plantPartnerUploadFile = document.getElementById('plant-partner-upload-file')
        plantPartnerUploadFile.addEventListener("dragover", this.fileDrag, false);
        plantPartnerUploadFile.addEventListener("dragleave", this.fileDrag, false);
        plantPartnerUploadFile.addEventListener("drop", (e) => this.plantPartnerUploadFileDrop(e), false);
    }

    fileDrag(event?: any) {
        event.stopPropagation();
        event.preventDefault();
        event.target.className = (event.type == ('dragover' || 'dragleave') ? 'drag-hover' : 'drop');
    }

    lowesLinkDailyRoutesFileDrop(event: any) {
        this.fileDrag(event);

        this.lowesLinkDailyRoutesFileName = event.dataTransfer.files[0].name

        const reader: FileReader = new FileReader();

        reader.onload = (e: any) => {
            this.lowesLinkData = [];
            this.dailyRoutesData = [];

            const bstr: string = e.target.result;
            const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

            this.lowesLinkData = <AOA>(XLSX.utils.sheet_to_json(
                wb.Sheets[wb.SheetNames[0]],
                {
                    header: [
                        'poNumber',
                        'storeNumber',
                        'storeDescription',
                        'itemNumber',
                        'itemDescription',
                        'quantityOrdered',
                        'quantityReceived'
                    ],
                    raw: false,
                })
            );

            this.dailyRoutesData = <AOA>(
                XLSX.utils.sheet_to_json(
                    wb.Sheets[wb.SheetNames[1]],
                    {
                        header: [
                            'loadNumber',
                            'vehicleCode',
                            'shipWeek',
                        ],
                        raw: false,
                    }
                )
                .map((route: any) => Object.assign({}, route, {
                    vehicleCode: route.vehicleCode ? Number(route.vehicleCode) : '',
                    shipWeek: route.shipWeek ? new Date(route.shipWeek).toLocaleDateString() : ''
                }))
            );

            if (this.plantPartnerData.length) {
                this.complete();
            }
        };

        reader.readAsBinaryString(event.dataTransfer.files[0]);
    }

    plantPartnerFileDrop(event: any) {
        this.fileDrag(event);

        this.plantPartnerFileName = event.dataTransfer.files[0].name

        const reader: FileReader = new FileReader();

        reader.onload = (e: any) => {
            this.plantPartnerData = [];

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(e.target.result, "text/xml");

            const sectionNodes = xmlDoc.getElementsByTagName("section");

            let store: any = { poNumber: '', shipDate: '', shipWeek: '', vehicleCode: '' };

            for (let i = 0; i < sectionNodes.length; i++) {
                const secitonTextNodes = sectionNodes[i].getElementsByTagName('text');

                if (secitonTextNodes.length) {

                    for (let j = 0; j < secitonTextNodes.length; j++) {
                        if (secitonTextNodes[j].attributes && secitonTextNodes[j].attributes.length) {
                            if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_vehicle_load_sheet.ship_date') {
                                if (store.poNumber) {
                                    this.plantPartnerData.push(store);
                                    store = { poNumber: '', shipDate: '', shipWeek: '', vehicleCode: '' };
                                }
                                if (secitonTextNodes[j].textContent.trim()) {
                                    const shipDateForWeek = new Date(secitonTextNodes[j].textContent.trim());
                                    const day = shipDateForWeek.getDay();
                                    const diff = shipDateForWeek.getDate() - day + (day == 0 ? -6 : 1);
                                    store.shipWeek = new Date(shipDateForWeek.setDate(diff)).toLocaleDateString();

                                    const shipDateForDate = new Date(secitonTextNodes[j].textContent.trim());
                                    store.shipDate = shipDateForDate.toLocaleDateString();
                                }
                            } else if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_vehicle_load_sheet.purchase_order_number') {
                                store.poNumber = (secitonTextNodes[j].textContent || '').trim();
                            } else if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_vehicle_load_sheet.vehicle_code') {
                                store.vehicleCode = secitonTextNodes[j].textContent ? Number((secitonTextNodes[j].textContent || '').trim()) : '';
                            } else if (secitonTextNodes[j].attributes[0].nodeValue === '@ItemPlusSku') {
                                const itemPlusSku = (secitonTextNodes[j].textContent || '').trim().split('*');
                                if (itemPlusSku.length && itemPlusSku[0] && itemPlusSku[1]) {
                                    store[itemPlusSku[1]] = itemPlusSku[0];
                                }
                            }
                        }
                    }
                }
            }

            if (store.poNumber) {
                this.plantPartnerData.push(store);
            }

            if (this.lowesLinkData.length) {
                this.complete();
            }
        };

        reader.readAsText(event.dataTransfer.files[0]);
    }

    complete() {
        const rows = [
            [
                'Load #',
                'Ship Week',
                'Ship Date',
                'Date Reported',
                'Purchase Order Number ID',
                'Location ID',
                'Location DESC',
                'Item Number ID',
                'Dewar Number ID',
                'Item Number DESC',
                'Quantity Ordered (Last Revised)',
                'Quantity Received'
            ]
        ].concat(
            this.lowesLinkData
            .map((store: any) => {
                const plantPartnerData = this.plantPartnerData.filter((order: any) => order.poNumber === store.poNumber);

                return Object.assign({}, store, plantPartnerData.length ? plantPartnerData[0] : {});
            })
            .map((store: any) => {
                const dailyRoutesData = this.dailyRoutesData.filter((route: any) => (route.shipWeek === store.shipWeek && route.vehicleCode === store.vehicleCode))
    
                return Object.assign(
                    {},
                    store,
                    {
                        loadNumber: dailyRoutesData && dailyRoutesData.length ? dailyRoutesData[0].loadNumber || '' : ''
                    }
                )
            })
            .map((store: any) => [
                store.loadNumber || '',
                store.shipWeek || '',
                store.shipDate || '',
                '',
                store.poNumber || '',
                store.storeNumber || '',
                store.storeDescription || '',
                store.itemNumber || '',
                store[store.itemNumber] || '',
                store.itemDescription || '',
                store.quantityOrdered || 0,
                store.quantityReceived || 0,
            ])
        );

        const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(rows);
        const writeWb: XLSX.WorkBook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(writeWb, ws, 'Sheet1');
        XLSX.writeFile(writeWb, 'Generated Lowes Link Report' + '.xlsx');
    }

    plantPartnerUploadFileDrop(event: any) {
        this.fileDrag(event);

        const reader: FileReader = new FileReader();

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const readWb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const rows = <AOA>(XLSX.utils.sheet_to_json(readWb.Sheets[readWb.SheetNames[0]], {
                header: [
                    'loadNumber',
                    'shipWeek',
                    'shipDate',
                    'dateReported',
                    'poNumber',
                    'storeNumber',
                    'locationDescription',
                    'itemNumber',
                    'dewarItemNumber',
                    'itemDescription',
                    'qtyOrdered',
                    'qtyReceived'
                ],
            })).map((row: any) => [
                'Order',
                'Edit',
                row.storeNumber || '',
                50996,
                row.poNumber || '',
                row.shipDate || '',
                row.dewarItemNumber || '',
                'I',
                row.qtyReceived || ''
            ]);

            rows.splice(0, 1, [
                'Type',
                'Action',
                'Store #',
                'Vendor #',
                'PO #',
                'Ship Date',
                'Product Code',
                'Product Code Type',
                'Import Qty'
            ]);

            const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
            const writeWb: XLSX.WorkBook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(writeWb, ws, 'Sheet1');
            XLSX.writeFile(writeWb, 'plant_partner_upload' + '.xls');
        };

        reader.readAsBinaryString(event.dataTransfer.files[0]);
    }

    showLowesLinkTemplate(template: string) {
        this.matDialog.open(TemplateDialogComponent,
            {
                minWidth: '250px',
                data: { template: template }
            });
    }
}
