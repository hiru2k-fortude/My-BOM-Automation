import { HttpClient } from '@angular/common/http';
import { Component } from '@angular/core';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.css'],
})
export class LoginComponent {
  inputValue: string = ''; // Variable to store input value
  inputValueEmail: string = ''; // Variable to store input value
  inputValuePassword: string = ''; // Variable to store input value
  apiData: any; // Variable to store API response

  message: string = ''; // Variable to store static message

  constructor(private http: HttpClient) {}

  onInputChange(field: string) {
    switch (field) {
      case 'name':
        console.log('Input Value (Name):', this.inputValue);
        break;
      case 'email':
        console.log('Input Value (Email):', this.inputValueEmail);
        break;
      case 'password':
        console.log('Input Value (Password):', this.inputValuePassword);
        break;
      default:
        break;
    }
  }

  makeApiRequest() {
    if (
      this.inputValue != '' &&
      this.inputValueEmail != '' &&
      this.inputValuePassword != ''
    ) {
      const apiUrl = 'http://localhost:3000/users'; // Example API URL

      const postData = {
        name: this.inputValue,
        email: this.inputValueEmail,
        password: this.inputValuePassword,
      };

      this.http.post(apiUrl, postData).subscribe(
        (data: any) => {
          console.log('POST Request Response:', data);
          this.apiData = data;
          this.message = ' - Data Inserted Successfully - ';
          alert(this.message);
          this.disableMessage();
          this.resetForm();
        },
        (error) => {
          this.message = 'Error Occurred while Inserting Data';
          this.disableMessage();
        }
      );
    } else {
      this.message = '-> Please Fill All The Fields <-';
      this.disableMessage();
    }
  }

  resetForm() {
    this.inputValue = '';
    this.inputValueEmail = '';
    this.inputValuePassword = '';
  }

  // Disable the message after 1 seconds
  disableMessage() {
    setTimeout(() => {
      this.message = '';
    }, 1000);
  }
}
