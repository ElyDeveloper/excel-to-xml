import { ApplicationConfig, provideZoneChangeDetection } from '@angular/core';
import { provideToastr } from 'ngx-toastr'; // Importar provideToastr

export const appConfig: ApplicationConfig = {
  providers: [
    provideZoneChangeDetection({ eventCoalescing: true }),
    provideToastr({
      timeOut: 3000,
      positionClass: 'toast-top-right',
      preventDuplicates: true,
      closeButton: true,
      progressBar: true,
    }),
  ],
};
