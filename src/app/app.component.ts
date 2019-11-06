import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
})
export class AppComponent {
    lowesLinkFile: any;
    lowesLinkData: any = [];

    plantPartnerFile: any;
    plantPartnerData: any = [];

    dailyRoutesFile: any;
    dailyRoutesData: any = [];

    plantPartnerFileUpload: any;

    wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };

    lowesLinkFileChange() {
        const ref: any = document.getElementById('lowes-link-file');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const stores = <AOA>(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }));

            stores
            .forEach(store =>
                this.lowesLinkData.push(
                    Object.assign({}, {
                        poNumber: store[0] ? store[0].toString() : '',
                        storeNumber: store[1] || '',
                        storeDescription: store[2] || '',
                        itemNumber: store[3] || '',
                        itemDescription: store[4] || '',
                        quantityOrdered: store[5] || '',
                        quantityReceived: store[6] || '',
                    })
                )
            )

            if (this.plantPartnerData.length && this.dailyRoutesData.length) {
                this.complete();
            }
        };

        reader.readAsBinaryString(target.files[0]);
    }

    plantPartnerFileChange() {
        const ref: any = document.getElementById('route-schedule-file');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(e.target.result, "text/xml");

            const sectionNodes = xmlDoc.getElementsByTagName("section");
            for (let i = 0; i < sectionNodes.length; i++) {
                const secitonTextNodes = sectionNodes[i].getElementsByTagName('text');

                if (secitonTextNodes.length) {
                    const store: any = { poNumber: '', shipDate: '', shipWeek: '', vehicleNumber: '' };

                    for (let j = 0; j < secitonTextNodes.length; j++) {
                        if (secitonTextNodes[j].attributes && secitonTextNodes[j].attributes.length) {
                            if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_delivery_route.po_number') {
                                store.poNumber = (secitonTextNodes[j].textContent || '').trim();
                            } else if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_delivery_route.ship_date') {
                                if (secitonTextNodes[j].textContent.trim()) {
                                    const shipDateForWeek = new Date(secitonTextNodes[j].textContent.trim());
                                    const day = shipDateForWeek.getDay();
                                    const diff = shipDateForWeek.getDate() - day + (day == 0 ? -6 : 1);
                                    const shipWeek = new Date(shipDateForWeek.setDate(diff)).toLocaleDateString();
                                    store.shipWeek = `${ shipWeek.slice(0, -4) }${ shipWeek.slice(-2) }`;

                                    const shipDateForDate = new Date(secitonTextNodes[j].textContent.trim());
                                    store.shipDate = `${ shipDateForDate.toLocaleDateString().slice(0, -4) }${ shipDateForDate.toLocaleDateString().slice(-2) }`;
                                }
                            } else if (secitonTextNodes[j].attributes[0].nodeValue === 'rpt_delivery_route.vehicle_code') {
                                store.vehicleNumber = secitonTextNodes[j].textContent ? Number((secitonTextNodes[j].textContent || '').trim()) : '';
                            }
                        }
                    }

                    if (store.poNumber) {
                        this.plantPartnerData.push(store);
                    }
                }
            }

            if (this.lowesLinkData.length && this.dailyRoutesData.length) {
                this.complete();
            }
        };

        reader.readAsText(target.files[0]);
    }

    dailyRoutesChange() {
        const ref: any = document.getElementById('daily-routes-file');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const routes = <AOA>(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }));

            routes
            .filter(route => route.length && route[0] && route[1] && route[2])
            .forEach(route => this.dailyRoutesData.push(Object.assign({}, { loadNumber: route[1], shipWeek: route[0], vehicleNumber: Number(route[2])})))

            if (this.lowesLinkData.length && this.plantPartnerData.length) {
                this.complete();
            }
        };

        reader.readAsBinaryString(target.files[0]);
    }

    complete() {
        const plantPartnerDailyRoutesMerge = this.plantPartnerData.map((store: any) => {
            const dailyRoutesData = this.dailyRoutesData.filter((route: any) => route.shipWeek === store.shipWeek && route.vehicleNumber === store.vehicleNumber);

            return Object.assign({}, store, {
                loadNumber: dailyRoutesData.length && dailyRoutesData[0] && dailyRoutesData[0].loadNumber ? dailyRoutesData[0].loadNumber : ''
            })
        });

        const lowesLinkMerge = this.lowesLinkData.map((store: any) => {
            const plantPartnerDailyRoutesData = plantPartnerDailyRoutesMerge.filter((route: any) => route.poNumber === store.poNumber);

            return Object.assign({}, store, plantPartnerDailyRoutesData.length ? plantPartnerDailyRoutesData[0] : {});
        });

        const rows = [
            [
                'Load #',
                'Ship Week',
                'Ship Date',
                'Date Reported',
                'PO #',
                'Store #',
                'Location Description',
                'Item #',
                'Item Description',
                'Qty Ordered',
                'Qty Received'
            ]
        ].concat(
            lowesLinkMerge.map((store: any) => [
                store.loadNumber || '',
                store.shipWeek || '',
                store.shipDate || '',
                '',
                store.poNumber || '',
                store.storeNumber || '',
                store.storeDescription || '',
                store.itemNumber || '',
                store.itemDescription || '',
                store.quantityOrdered || 0,
                store.quantityReceived || 0,
            ])
        );

        const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(rows);
        const writeWb: XLSX.WorkBook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(writeWb, ws, 'Sheet1');
        XLSX.writeFile(writeWb, 'report' + '.xlsx');
    }

    plantPartnerFileUploadChange() {
        const ref: any = document.getElementById('plant-partner-file-upload');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const readWb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const rows = <AOA>(XLSX.utils.sheet_to_json(readWb.Sheets[readWb.SheetNames[0]], { header: 1 }))
            .map((store: any) => [
                'Order',
                'Edit',
                store[5] || '',
                50996,
                store[4] || '',
                store[2] ? new Date(store[2]).toLocaleDateString() : '',
                // store[7] || '', ???
                '',
                'I',
                store[10] || 0
            ])

            rows.shift();

            const newRows = [
                [
                    'Type',
                    'Action',
                    'Store #',
                    'Vendor #',
                    'PO #',
                    'Ship Date',
                    'Product Code',
                    'Product Code Type',
                    'Import Qty',
                ]
            ].concat(rows);

            const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(newRows);
            const writeWb: XLSX.WorkBook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(writeWb, ws, 'Sheet1');
            XLSX.writeFile(writeWb, 'plant_partner_upload' + '.xls');
        };

        reader.readAsBinaryString(target.files[0]);
    }
}
