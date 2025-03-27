import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelToXmlComponent } from './excel-to-xml.component';

describe('ExcelToXmlComponent', () => {
  let component: ExcelToXmlComponent;
  let fixture: ComponentFixture<ExcelToXmlComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExcelToXmlComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExcelToXmlComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
