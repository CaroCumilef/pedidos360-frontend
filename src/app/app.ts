import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpHeaders, HttpClientModule } from '@angular/common/http';
import { MsalService, MsalModule } from '@azure/msal-angular';
import { AuthenticationResult } from '@azure/msal-browser';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, HttpClientModule, MsalModule],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'Pedidos360 - Frontend';
  usuario: string = '';
  productos: any[] = [];
  apiGatewayUrl = 'https://i5xfqp6s96.execute-api.us-east-1.amazonaws.com/productos';

  constructor(
    private authService: MsalService,
    private http: HttpClient
  ) {}

  ngOnInit(): void {
    // Verificar si el usuario ya inició sesión con Azure AD
    const accounts = this.authService.instance.getAllAccounts();
    if (accounts.length > 0) {
      this.authService.instance.setActiveAccount(accounts[0]);
      this.usuario = accounts[0].name || accounts[0].username;
    }
  }

  // Método para Iniciar Sesión con Azure AD / Entra ID
  iniciarSesion(): void {
    this.authService.loginPopup({
      scopes: ['openid', 'profile', 'user.read']
    }).subscribe({
      next: (result: AuthenticationResult) => {
        this.authService.instance.setActiveAccount(result.account);
        this.usuario = result.account.name || result.account.username;
        console.log('Sesión iniciada con éxito:', result);
      },
      error: (error) => console.error('Error al iniciar sesión:', error)
    });
  }

  // Método para Cerrar Sesión
  cerrarSesion(): void {
    this.authService.logoutPopup();
    this.usuario = '';
    this.productos = [];
  }

  // Método para invocar el API Gateway con Token JWT (Respuesta 200 OK)
  obtenerProductos(): void {
    const account = this.authService.instance.getActiveAccount();

    if (!account) {
      console.warn('No hay usuario autenticado.');
      return;
    }

    // Adquirir el token de forma silenciosa para adjuntarlo a la petición HTTP
    this.authService.acquireTokenSilent({
      account: account,
      scopes: ['openid', 'profile', 'user.read']
    }).subscribe({
      next: (response: AuthenticationResult) => {
        // Se utiliza idToken prioritariamente para validar claims en API Gateway
        const token = response.idToken || response.accessToken;

        // Construir la cabecera Authorization con el token Bearer
        const headers = new HttpHeaders({
          'Authorization': `Bearer ${token}`
        });

        // Llamada al API Gateway con los headers
        this.http.get(this.apiGatewayUrl, { headers }).subscribe({
          next: (data: any) => {
            this.productos = data;
            console.log('Productos recibidos exitosamente (HTTP 200 OK):', data);
          },
          error: (err) => console.error('Error al invocar API Gateway:', err)
        });
      },
      error: (error) => {
        console.error('Error obteniendo el token de acceso:', error);
      }
    });
  }
}