page 60100 "MS Teams Setup"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "MS Teams Setup";

    layout
    {
        area(Content)
        {
            group(Access)
            {
                Caption = 'App registration';

                field("Client ID"; ClientID)
                {
                    ApplicationArea = All;
                    ExtendedDatatype = Masked;
                    Tooltip = 'Specifies the application client ID for the Azure App Registration that accesses the storage account.';

                    trigger OnValidate()
                    begin
                        MsTeamsCredentials.SetClientID(ClientID);
                    end;
                }
                field("Client Secret"; ClientSecret)
                {
                    ApplicationArea = All;
                    ExtendedDatatype = Masked;
                    Tooltip = 'Specifies the client secret for the Azure App Registration that accesses the storage account.';

                    trigger OnValidate()
                    begin
                        MsTeamsCredentials.SetClientSecret(ClientSecret);
                    end;
                }
                field(OAuthAuthorityUrl; rec."OAuth Authority Url")
                {
                    ApplicationArea = All;
                    Tooltip = 'Specifies the Oauth Authority Url that can be found on the app registration.';

                }
                field(MSTeamsRootQueryURL; rec."MS Teams Root Query URL")
                {
                    ApplicationArea = All;
                    Tooltip = 'The base URL of the connection with MS graph.';

                }
                field(RedirectUrl; rec."Redirect URL")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the redirect URL of the application that will be used to to the MS Teams integration.';

                }
            }
        }
    }

    var
        [NonDebuggable]
        ClientID: Text;
        [NonDebuggable]
        ClientSecret: Text;
        MsTeamsCredentials: Codeunit "MS Teams Credentials";

    trigger OnOpenPage()
    begin
        if not Rec.Get() then begin
            Rec.Init();
            InitializeDefaultRedirectUrl();
            rec."MS Teams Root Query URL" := 'https://graph.microsoft.com/v1.0';
            Rec.Insert();
        end else begin
            MsTeamsCredentials.Init();
            ClientID := MsTeamsCredentials.GetClientID();
            ClientSecret := MsTeamsCredentials.GetClientSecret();
        end;
    end;

    local procedure InitializeDefaultRedirectUrl()
    var
        OAuth2: Codeunit "OAuth2";
        RedirectUrl: Text;
    begin
        OAuth2.GetDefaultRedirectUrl(RedirectUrl);
        Rec."Redirect URL" := CopyStr(RedirectUrl, 1, MaxStrLen(Rec."Redirect URL"));
    end;
}