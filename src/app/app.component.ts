import { Component } from '@angular/core';
// import { ExcelToXmlComponent } from './excel-to-xml/excel-to-xml.component';
import { ExcelViewerComponent } from './excel-viewer/excel-viewer.component';

@Component({
  selector: 'app-root',
  standalone:true,
  imports: [ExcelViewerComponent],
  template: `
    <div class="app-container">
      <header class="app-header">
        <h1>Conversor Excel a XML</h1>
      </header>
      <main>
        <app-excel-viewer></app-excel-viewer>
        <!-- <app-excel-to-xml></app-excel-to-xml> -->
      </main>
      <footer class="app-footer">
        <p>© 2025 - Conversor Excel a XML</p>
      </footer>
    </div>
  `,
  styles: [`
    .app-container {
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }
    
    .app-header {
      background-color: #3366cc;
      color: white;
      padding: 1rem;
      text-align: center;
    }
    
    main {
      flex-grow: 1;
      padding: 1rem 0;
    }
    
    .app-footer {
      background-color: #f8f9fa;
      padding: 1rem;
      text-align: center;
      font-size: 0.875rem;
      color: #6c757d;
    }
  `]
})
export class AppComponent {
  title = 'excel-to-xml';
}