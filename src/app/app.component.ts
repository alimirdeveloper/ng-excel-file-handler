import { Component, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { ExcelColumn, ExcelFileHandlerComponent } from '../components/excel-file-handler/excel-file-handler.component';
import { Observable, of, delay } from 'rxjs';
import { JsonPipe, PercentPipe } from '@angular/common';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    RouterOutlet,
    JsonPipe,
    PercentPipe,
    ExcelFileHandlerComponent
  ],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'ng-excel-file-handler';
  uploadedData = signal<any[]>([])
  items: any[] = [
    {
      id: '1',
      label: 'test',
    },
    {
      id: '2',
      label: 'test2',
      disabled: true
    },
    {
      id: '3',
      label: 'تست',
    },
    {
      id: '4',
      label: 'test4',
    }
  ]

  testObser = (): Observable<{ key: string; value: string }[]> => {
    return of(this.items.map(a => {
      const item = {
        key: a.id,
        value: a.label
      }
      return item
    })).pipe(
      delay(5000)
    )
  }

  excelHeaders: ExcelColumn[] = [

    {
      name: 'Name',
      idName: 'Name',
      type: 'text',
      mandatory: true,
    },
    {
      name: 'Age',
      idName: 'Age',
      type: 'number',

    },
    {
      name: 'Gender',
      idName: 'Gender',
      type: 'drop-down',
      options: [
        { key: '1', value: 'Male' },
        { key: '0', value: 'Female' },
      ]
    },
    {
      name: 'Family',
      idName: 'Family',
      type: 'text',

    },
    {
      name: 'Education',
      idName: 'Education',
      type: 'drop-down',
      options: [
        { key: '0', value: 'diploma' },
        { key: '1', value: "Bachelor's degree" },
        { key: '2', value: "Master's degree" },
        { key: '3', value: 'Ph.D' }
      ]
    },
    {
      name: 'stock',
      idName: 'stock',
      type: 'Percentage',

    },
    {
      name: 'observable Item',
      idName: 'observable Item',
      type: 'drop-down',
      optionsObeservable: this.testObser
    },
    {
      name:'id',
      idName:'id',
      type:'number',
      mandatory:true, 
      unique:true
    }
  ]
  onDataUploaded(excelData: any[]) {
    this.uploadedData.set(excelData || []);
  }
}
