import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelFileHandlerComponent } from './excel-file-handler.component';

describe('ExcelFileHandlerComponent', () => {
  let component: ExcelFileHandlerComponent;
  let fixture: ComponentFixture<ExcelFileHandlerComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExcelFileHandlerComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ExcelFileHandlerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
