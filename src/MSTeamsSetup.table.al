table 60100 "MS Teams Setup"
{
    DataClassification = CustomerContent;
    Access = Internal;
    DataPerCompany = true;

    fields
    {
        field(1; "Primary key"; Code[10])
        {
            Caption = 'Primary Key';
            Editable = false;
        }
        field(2; "OAuth Authority Url"; Text[250])
        {
            Caption = 'OAuth Authority Url';
        }
        field(3; "MS Teams Root Query URL"; Text[250])
        {
            Caption = 'MS Teams Root Query URL';
        }
        field(4; "Redirect URL"; Text[250])
        {
            Caption = 'Redirect URL';
        }
    }

    keys
    {
        key(Key1; "Primary key")
        {
            Clustered = true;
        }
    }

}