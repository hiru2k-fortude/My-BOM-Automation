import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor() {
    // Set a baseurl
    localStorage.setItem('UEE-BOM-Automation-backend-baseurl', 'http://localhost:7101');
  }
}
