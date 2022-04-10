permissionset 88000 EmailNewObject
{
    Assignable = true;
    Caption = 'EmailNewObject', MaxLength = 30;
    Permissions =
        table "Import and Send Attachment" = X,
        tabledata "Import and Send Attachment" = RMID,
        page "Import and Send Attachment" = X,
        report "Custom Email" = X;
}
