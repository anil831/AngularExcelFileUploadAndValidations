import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  file!: File;
  errorMessage!: string;

  onFileSelected(event: any) {
    this.file = event.target.files[0];
    this.validateFile();
  }

  validateFile() {
    if (!this.file) {
      this.errorMessage = 'Please select a file.';
      return;
    }

    if (!this.file.name.endsWith('.xlsx')) {
      this.errorMessage = 'Please select an Excel file.';
      return;
    }

    if (this.file.size > 1024 * 1024) {
      this.errorMessage = 'File size must be less than 1MB.';
      return;
    }

    const fileReader = new FileReader();
    fileReader.onload = (e) => {
      const arrayBuffer = fileReader.result as ArrayBuffer;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(worksheet) as { [key: string]: string | number }[];

      if (data.length === 0) {
        this.errorMessage = 'File is empty.';
        return;
      }

      const requiredFields = ['Entity', 'Sector', 'Nature', 'DOF Nature', 'Chapter', 'Location',
        'Location name', 'CC', 'CC Name', 'Account', 'Account Name', 'aa', 'ss', 'dfs', 'sdsas', 'Programme',
        'Programme name', 'bb', 'asw', 'Budget', ' Encumbrance', 'Actual', 'Funds Available', 'Financial Year'];
      const missingFields = requiredFields.filter(field => !Object.keys(data[0]).includes(field));
      if (missingFields.length > 0) {
        this.errorMessage = `Required field(s) ${missingFields.join(', ')} is/are missing.`;
        return;
      }

      const missingValues = data.filter((row) => {
        return requiredFields.some(field => {
          return !row.hasOwnProperty(field) || !row[field];
        });
      });
      if (missingValues.length > 0) {
        this.errorMessage = `Field value(s) missing in row(s) ${missingValues.map(row => data.indexOf(row) + 2).join(', ')}.`;
        return;
      }

      const uniqueData = new Set(data.map(row => JSON.stringify(row)));
      if (uniqueData.size !== data.length) {
        this.errorMessage = 'File contains duplicate rows.';
        return;
      }

      this.errorMessage = '';
      // Do something with the validated data
    };
    fileReader.readAsArrayBuffer(this.file);
  }
}
