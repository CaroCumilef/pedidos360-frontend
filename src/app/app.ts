import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { RouterOutlet } from '@angular/router';
import { MsalService } from '@azure/msal-angular';
import { AuthenticationResult } from '@azure/msal-browser';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class AppComponent implements OnInit {
  isLoggedIn = false;
  apiData: any = null;
  apiResponse: any = null;
  errorMsg: string = '';
  userName: string = '';

  constructor(
    private msalService: MsalService, 
    private http: HttpClient
  ) {}

  ngOnInit(): void {
    this.msalService.handleRedirectObservable().subscribe({
      next: (result: AuthenticationResult | null) => {
        if (result && result.account) {
          this.msalService.instance.setActiveAccount(result.account);
        }
        this.checkAccount();
      },
      error: (error) => console.error('Error en el callback de MSAL:', error)
    });

    this.checkAccount();
  }

  checkAccount(): void {
    const activeAccount = this.msalService.instance.getActiveAccount() || 
                          this.msalService.instance.getAllAccounts()[0];
    
    if (activeAccount) {
      this.msalService.instance.setActiveAccount(activeAccount);
      this.isLoggedIn = true;
      this.userName = activeAccount.name || activeAccount.username;
    } else {
      this.isLoggedIn = false;
      this.userName = '';
    }
  }

  login(): void {
    this.msalService.loginRedirect();
  }

  logout(): void {
    this.msalService.logoutRedirect();
  }

  consultarApi(): void {
    this.errorMsg = '';
    this.apiData = null;
    this.apiResponse = null;

    const account = this.msalService.instance.getActiveAccount();

    if (!account) {
      this.errorMsg = 'No hay una cuenta activa de Azure AD.';
      return;
    }

    // 1. Obtener el token JWT en silencio
    this.msalService.acquireTokenSilent({
      scopes: ['openid', 'profile', 'email'],
      account: account
    }).subscribe({
      next: (response: AuthenticationResult) => {
        // ID Token contiene el emisor (iss) y audiencia configurados en API Gateway
        const token = response.idToken || response.accessToken;
        
        // 2. Adjuntar el token como Bearer Header
        const headers = new HttpHeaders({
          'Authorization': `Bearer ${token}`
        });

        // 3. Invocar API Gateway
        this.http.get('https://i5xfqp6s96.execute-api.us-east-1.amazonaws.com/api/pedidos', { headers })
          .subscribe({
            next: (data) => {
              this.apiData = data;
              this.apiResponse = data;
              console.log('Respuesta 200 OK desde AWS API Gateway:', data);
            },
            error: (err) => {
              this.errorMsg = JSON.stringify(err, null, 2);
              console.error('Error al invocar API Gateway:', err);
            }
          });
      },
      error: (err) => {
        console.error('Error al adquirir token:', err);
        this.errorMsg = 'No se pudo obtener el token de autenticación.';
      }
    });
  }
}