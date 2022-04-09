table 88000 "Import and Send Attachment"
{
    Caption = 'Import and Send Attachment';
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Code"; Code[20])
        {
            Caption = 'Code';
            DataClassification = CustomerContent;
        }
        field(5; "Send To"; Text[250])
        {
            DataClassification = CustomerContent;
        }
        

    }
    keys
    {
        key(PK; "Code")
        {
            Clustered = true;
        }
    }
    var
    
}
