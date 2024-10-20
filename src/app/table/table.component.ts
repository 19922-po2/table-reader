import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { HotTableModule } from '@handsontable/angular';
import { registerAllModules } from 'handsontable/registry';
import {
  registerPlugin, // plugins' registering function
  UndoRedo,
} from 'handsontable/plugins';
import { HotTableRegisterer } from '@handsontable/angular';
import Handsontable from 'handsontable/base';
import * as XLSX from 'xlsx';

type AOA = any[][];

// register Handsontable's modules
registerAllModules();
registerPlugin(UndoRedo);

@Component({
  selector: 'app-table',
  standalone: true,
  imports: [CommonModule, HotTableModule],
  templateUrl: './table.component.html',
  styleUrl: './table.component.css',
})
export class TableComponent {
  constructor() {}

  private hotRegisterer = new HotTableRegisterer();

  dataset: any[] = [
    { id: 1, country: 'Portugal', city: 'Lisbon' },
    { id: 2, country: 'Spain', city: 'Madrid' },
    { id: 3, country: 'France', city: 'Paris' },
    { id: 4, country: 'United Kingdom', city: 'London' },
  ];

  hotSettings: Handsontable.GridSettings = {
    data: this.dataset,
    colHeaders: ['ID', 'Country', 'City'],
    height: 'auto',
    autoWrapRow: true,
    autoWrapCol: true,
    licenseKey: 'non-commercial-and-evaluation',
  };
  hot: Handsontable | undefined;
  tableid = 'tableid';
  error = '';

  ngAfterViewInit() {
    //this.hot = document.getElementById('tableid');
    this.hot = new Handsontable(
      document.getElementById('tableid')!,
      this.hotSettings
    );

    console.log('ngOnInit Data:', this.hot);
  }

  printData() {
    console.log('Dataset:');
    console.log(this.dataset);
  }

  addRow() {
    this.hot!.alter('insert_row_below', this.hot!.countRows());
  }

  downloadExcell() {
    console.log('Dowloading...');
    const exportPlugin = this.hot!.getPlugin('exportFile');
    exportPlugin.downloadFile('csv', {
      bom: false,
      columnDelimiter: ',',
      columnHeaders: false,
      exportHiddenColumns: true,
      exportHiddenRows: true,
      fileExtension: 'csv',
      filename: 'test',
      mimeType: 'text/csv',
      rowDelimiter: '\r\n',
      rowHeaders: true,
    });

    // needs to be converted to xlsx
  }

  uploadExcell() {
    console.log('Uploading...');
  }

  data: AOA = [
    [1, 2],
    [3, 4],
  ];
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>XLSX.utils.sheet_to_json(ws, { header: 1 });
      console.log('data:', this.data);
      /* this.data.map((res) => {
        if (res[0] === 'no') {
          console.log(res[0]);
        } else {
          console.log(res[0]);
        }
      }); */
      console.log('data->', this.data);
      this.dataset = this.data;
      this.hot!.getInstance().loadData([]);
      this.hot!.getInstance().loadData(this.data);
      this.hot!.render();
    };
    reader.readAsBinaryString(target.files[0]);
  }
}
