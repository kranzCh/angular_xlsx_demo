import { PortfolioSummary } from './PortfolioSummary';
import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'xlsx_demo';
  rows: PortfolioSummary[] = [];
  data: any[] = [];
  fileName: string = 'SheetJS.xlsx';
  reader: FileReader = new FileReader();
  target: DataTransfer = new DataTransfer();

  onFileChanged(event: any) {
    /* wire up file reader */
    this.target = <DataTransfer>event.target;
    if (this.target.files.length !== 1)
      throw new Error('Cannot use multiple files');
    this.reader.onload = (e: any) => {
      const bstr = e.target.result;
      /* parse workbook */

      // const url = "https://sheetjs.com/data/PortfolioSummary.xls";
      // const workbook = XLSX.read(await (await fetch(url)).arrayBuffer());

      const workbook = XLSX.read(bstr, { type: 'binary' });

      /* get first worksheet */
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw_data = XLSX.utils.sheet_to_json<any>(worksheet, {
        header: 1,
      });

      /* fill years */
      let last_year = 0;
      raw_data.forEach(
        (r) => (last_year = r[0] = r[0] != null ? r[0] : last_year)
      );

      /* select data rows */
      const rows = raw_data.filter((r) => r[0] >= 2007 && r[0] <= 2023);

      /* generate row objects */
      this.rows = rows.map((r) => ({
        year: r[0],
        quarter: r[1],
        totalDO: r[8],
      }));
    };
    console.log(this.rows);
  }

  import() {
    this.reader.readAsBinaryString(this.target.files[0]);
  }

  export(): void {
    for (let row of this.rows) {
      this.data.push([row.year, row.quarter, row.totalDO]);
    }
    /* generate worksheet */
    const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const workbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    /* fix headers */
    XLSX.utils.sheet_add_aoa(worksheet, [['Fiscal Year', 'Quarter', 'Total']], {
      origin: 'A1',
    });

    /* save to file */
    XLSX.writeFile(workbook, this.fileName);
    console.log(workbook);
  }
}
