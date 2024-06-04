import { Injectable } from '@angular/core';
import { Observable, throwError } from 'rxjs';
import { catchError } from 'rxjs/operators';
import { HttpClient, HttpHeaders } from '@angular/common/http';

@Injectable({
  providedIn: 'root',
})
export class FileUploadService {
  private readonly baseUrl = localStorage.getItem(
    'UEE-BOM-Automation-backend-baseurl'
  );

  constructor(private http: HttpClient) {}

  uploadFile(file: File): Observable<any> {
    const formData = new FormData();
    formData.append('excelFile', file, file.name);

    return this.http.post<any>(`${this.baseUrl}/excel`, formData).pipe(
      catchError((error) => {
        return throwError(error);
      })
    );
  }

  // DeleteOldFiles(name: string): Observable<any> {
  //   return this.http.post<any>(`${this.baseUrl}/excel/DeleteOldFiles`, {
  //     name: name,
  //   });
  // }

  // FillTemplate(body: any, Style : string) {
  //   return this.http.post<any>(`${this.baseUrl}/excel/fillTemplate`, {
  //     List: body,
  //     FileName : localStorage.getItem('UEE-BOM-Automation-backend-file-name'),
  //     Style : Style
  //   });
  // }
}
