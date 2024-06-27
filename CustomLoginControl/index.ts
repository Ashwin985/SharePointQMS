import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { PublicClientApplication, AuthenticationResult } from "@azure/msal-browser";

export class CustomLoginControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _value: string;
    private _notifyOutputChanged: () => void;
    private buttonElement: HTMLButtonElement;
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private msalConfig = {
        auth: {
            clientId: "YOUR_CLIENT_ID", // Replace with your Azure AD app client ID
            authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Replace with your Azure AD tenant ID
            redirectUri: "YOUR_REDIRECT_URI" // Replace with your redirect URI
        }
    };
    private msalInstance: PublicClientApplication;

    constructor() {
        this.msalInstance = new PublicClientApplication(this.msalConfig);
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = document.createElement("div");
        this._notifyOutputChanged = notifyOutputChanged;

        // Create the login button
        this.buttonElement = document.createElement("button");
        this.buttonElement.innerText = "Login";
        this.buttonElement.addEventListener("click", this.login.bind(this));

        // Append the button to the container
        this._container.appendChild(this.buttonElement);
        container.appendChild(this._container);
    }

    private async login(): Promise<void> {
        try {
            const loginResponse: AuthenticationResult = await this.msalInstance.loginPopup();
            this._value = loginResponse.account.username;
            this._notifyOutputChanged();
            this.callGraphAPI(loginResponse.accessToken);
        } catch (error) {
            console.error("Login failed", error);
        }
    }

    private async callGraphAPI(accessToken: string): Promise<void> {
        try {
            const response = await fetch("https://graph.microsoft.com/v1.0/me", {
                headers: {
                    Authorization: `Bearer ${accessToken}`
                }
            });
            const data = await response.json();
            console.log(data); // Handle the data as needed
        } catch (error) {
            console.error("Error calling Graph API", error);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
    }

    public getOutputs(): IOutputs {
        return {
            sampleProperty: this._value
        };
    }

    public destroy(): void {
        this.buttonElement.removeEventListener("click", this.login.bind(this));
    }
}
