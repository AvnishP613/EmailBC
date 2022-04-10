page 88000 "Import and Send Attachment"
{
    ApplicationArea = All;
    Caption = 'Import and Send Attachment';
    PageType = List;
    SourceTable = "Import and Send Attachment";
    UsageCategory = Lists;

    layout
    {
        area(content)
        {
            repeater(General)
            {
                field("Code"; Rec."Code")
                {
                    ToolTip = 'Specifies the value of the Code field.';
                    ApplicationArea = All;
                }
                field("Send To"; Rec."Send To")
                {
                    ToolTip = 'Specifies the value of the Send To field.';
                    ApplicationArea = All;
                }

            }
        }



    }
    actions
    {
        area(Processing)
        {
            action(EmailItem)
            {
                ApplicationArea = All;
                Caption = 'Email Item';
                Image = SendEmailPDF;
                PromotedCategory = Category4;
                Promoted = true;
                ToolTip = 'Email Item.';

                trigger OnAction()
                var
                    EmailItem: Record "Email Item";
                    OutStr: OutStream;
                    bODYtEST: Text;
                    TempBlob: Codeunit "Temp Blob";
                    outStreamReport: OutStream;
                    inStreamReport: InStream;
                    TempBlobArray: Array[10] of Codeunit "Temp Blob";
                    ArrayOutReport: Array[10] of OutStream;
                    ArrayInstrReport: Array[10] of InStream;
                    i: Integer;
                    RecRef: RecordRef;
                    SHIPMENT: Record "Sales Shipment Header";
                    SInvH: Record "Sales Invoice Header";
                    FullFileName: Text;

                    OutS: OutStream;
                    InS: InStream;
                    B: Codeunit "Temp Blob";
                    Ref: RecordRef;
                    FRef: FieldRef;
                    Recipients: List of [Text];
                    Body: Text;
                    FileMgmt: Codeunit "File Management";
                    FldRef: FieldRef;
                begin

                    Rec.TestField("Code");
                    Rec.TestField("Send To");
                    bODYtEST := 'Hi,' + '<br><br>' + 'Email Item';

                    EmailItem.Initialize();
                    EmailItem."Send to" := Rec."Send To";
                    EmailItem.Subject := 'Email Item';
                    EmailItem.Validate("Plaintext Formatted", false);
                    EmailItem.Validate("Send as HTML", true);
                    EmailItem.SetBodyText(bODYtEST);


                    IF SHIPMENT.FindFirst() then;
                    clear(RecRef);
                    RecRef.GetTable(SHIPMENT);

                    FldRef := RecRef.Field(SHIPMENT.FieldNo("No."));
                    FldRef.SetRange(SHIPMENT."No.");
                    if RecRef.FindFirst() then begin

                        //Generate blob from report
                        TempBlob.CreateOutStream(outStreamReport);
                        TempBlob.CreateInStream(inStreamReport);
                        Report.SaveAs(Report::"Sales - Shipment", '', ReportFormat::Pdf, outStreamReport, RecRef);
                        EmailItem.AddAttachment(inStreamReport, 'Shipment.pdf');
                    end;
                    i := 1;


                    if SInvH.FindSet() then
                        repeat

                            TempBlobArray[i].CreateOutStream(ArrayOutReport[i]);
                            TempBlobArray[i].CreateInStream(ArrayInstrReport[i]);
                            FullFileName := 'Invoice' + SInvH."No." + '.' + 'pdf';
                            Clear(FldRef);
                            clear(RecRef);
                            RecRef.GetTable(SInvH);
                            FldRef := RecRef.Field(SInvH.FieldNo("No."));
                            FldRef.SetRange(SInvH."No.");
                            if RecRef.FindFirst() then begin

                                Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, ArrayOutReport[i], RecRef);

                                EmailItem.AddAttachment(ArrayInstrReport[i], FullFileName);
                            end;
                            i += 1;

                        until (SInvH.Next = 0) or (i = 3);



                    EmailItem.Send(true, Enum::"Email Scenario"::Default);

                end;
            }

            action(EmailItemReportBody)
            {
                ApplicationArea = All;
                Caption = 'Email Item Report Body';
                Image = SendEmailPDF;
                PromotedCategory = Category4;
                Promoted = true;
                ToolTip = 'Email Item.';

                trigger OnAction()
                var
                    EmailItem: Record "Email Item";
                    OutStr: OutStream;
                    bODYtEST: Text;
                    TempBlob: Codeunit "Temp Blob";
                    outStreamReport: OutStream;
                    inStreamReport: InStream;
                    TempBlobArray: Array[10] of Codeunit "Temp Blob";
                    ArrayOutReport: Array[10] of OutStream;
                    ArrayInstrReport: Array[10] of InStream;
                    i: Integer;
                    RecRef: RecordRef;
                    SHIPMENT: Record "Sales Shipment Header";
                    SInvH: Record "Sales Invoice Header";
                    FullFileName: Text;

                    OutS: OutStream;
                    InS: InStream;
                    Ref: RecordRef;
                    FRef: FieldRef;
                    Recipients: List of [Text];
                    Body: Text;
                    FileMgmt: Codeunit "File Management";
                    FldRef: FieldRef;
                begin

                    Rec.TestField("Code");
                    Rec.TestField("Send To");

                    EmailItem.Initialize();
                    EmailItem."Send to" := Rec."Send To";
                    EmailItem.Subject := 'Email Item';
                    EmailItem.Validate("Plaintext Formatted", false);
                    EmailItem.Validate("Send as HTML", true);

                    //Body
                    clear(RecRef);
                    TempBlob.CreateOutStream(outStreamReport);
                    Report.SaveAs(Report::"Custom Email", '', ReportFormat::Html, outStreamReport, RecRef);
                    TempBlob.CreateInStream(inStreamReport);
                    inStreamReport.ReadText(Body);
                    EmailItem.SetBodyText(Body);
                    //Body


                    //Attachments
                    i := 1;


                    if SInvH.FindSet() then
                        repeat

                            TempBlobArray[i].CreateOutStream(ArrayOutReport[i]);
                            TempBlobArray[i].CreateInStream(ArrayInstrReport[i]);
                            FullFileName := 'Invoice' + SInvH."No." + '.' + 'pdf';
                            Clear(FldRef);
                            clear(RecRef);
                            RecRef.GetTable(SInvH);
                            FldRef := RecRef.Field(SInvH.FieldNo("No."));
                            FldRef.SetRange(SInvH."No.");
                            if RecRef.FindFirst() then begin

                                Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, ArrayOutReport[i], RecRef);

                                EmailItem.AddAttachment(ArrayInstrReport[i], FullFileName);
                            end;
                            i += 1;

                        until (SInvH.Next = 0) or (i = 3);
                    //Attachments


                    EmailItem.Send(true, Enum::"Email Scenario"::Default);

                end;
            }


            action(EmailAcc)
            {
                ApplicationArea = All;
                Caption = 'Email Account';
                Image = Email;
                PromotedCategory = Category10;
                Promoted = true;
                ToolTip = 'Email Account.';

                trigger OnAction()
                var
                    EmailMessage: Codeunit "Email Message";
                    Email: Codeunit Email;
                    bODYtEST: Text;
                    TempBlob: Codeunit "Temp Blob";
                    outStreamReport: OutStream;
                    inStreamReport: InStream;
                    TempBlobArray: Array[10] of Codeunit "Temp Blob";
                    ArrayOutReport: Array[10] of OutStream;
                    ArrayInstrReport: Array[10] of InStream;
                    i: Integer;
                    RecRef: RecordRef;
                    SHIPMENT: Record "Sales Shipment Header";
                    AttachmentName: Text;
                    SInvH: Record "Sales Invoice Header";
                    FullFileName: Text;
                    FldRef: FieldRef;



                begin
                    Rec.TestField("Code");
                    Rec.TestField("Send To");

                    bODYtEST := 'Hi,' + '<br><br>' + 'Email Account';

                    EmailMessage.Create(Rec."Send To", 'Email Account', bODYtEST, true);

                    IF SHIPMENT.FindFirst() then;
                    clear(RecRef);

                    RecRef.GetTable(SHIPMENT);
                    FldRef := RecRef.Field(SHIPMENT.FieldNo("No."));
                    FldRef.SetRange(SHIPMENT."No.");

                    //Generate blob from report
                    TempBlob.CreateOutStream(outStreamReport);
                    TempBlob.CreateInStream(inStreamReport);
                    Report.SaveAs(Report::"Sales - Shipment", '', ReportFormat::Pdf, outStreamReport, RecRef);
                    AttachmentName := 'Shipment.pdf';
                    EmailMessage.AddAttachment(AttachmentName, 'Shipment', inStreamReport);

                    i := 1;


                    if SInvH.FindSet() then
                        repeat

                            TempBlobArray[i].CreateOutStream(ArrayOutReport[i]);
                            TempBlobArray[i].CreateInStream(ArrayInstrReport[i]);
                            FullFileName := 'Invoice' + SInvH."No." + '.' + 'pdf';
                            clear(RecRef);
                            RecRef.GetTable(SInvH);
                            FldRef := RecRef.Field(SInvH.FieldNo("No."));
                            FldRef.SetRange(SInvH."No.");
                            if RecRef.FindFirst() then begin

                                Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, ArrayOutReport[i], RecRef);

                                EmailMessage.AddAttachment(FullFileName, 'Invoice', ArrayInstrReport[i]);
                            end;
                            i += 1;

                        until (SInvH.Next = 0) or (i = 3);


                    if not Email.Send(EmailMessage, Enum::"Email Scenario"::Default) then begin
                        Message(GetLastErrorText);
                        ClearLastError();

                    end else
                        Message('Success');

                end;
            }

            action(EmailAccHTMLBody)
            {
                ApplicationArea = All;
                Caption = 'Email Account HTML Body';
                Image = Email;
                PromotedCategory = Category10;
                Promoted = true;
                ToolTip = 'Email Account.';

                trigger OnAction()
                var
                    EmailMessage: Codeunit "Email Message";
                    Email: Codeunit Email;
                    TempBlob: Codeunit "Temp Blob";
                    outStreamReport: OutStream;
                    inStreamReport: InStream;
                    TempBlobArray: Array[10] of Codeunit "Temp Blob";
                    ArrayOutReport: Array[10] of OutStream;
                    ArrayInstrReport: Array[10] of InStream;
                    i: Integer;
                    RecRef: RecordRef;
                    SHIPMENT: Record "Sales Shipment Header";
                    AttachmentName: Text;
                    SInvH: Record "Sales Invoice Header";
                    FullFileName: Text;
                    FldRef: FieldRef;

                    Body: Text;

                begin
                    Rec.TestField("Code");
                    Rec.TestField("Send To");


                    //Body
                    clear(RecRef);
                    TempBlob.CreateOutStream(outStreamReport);
                    Report.SaveAs(Report::"Custom Email", '', ReportFormat::Html, outStreamReport, RecRef);
                    TempBlob.CreateInStream(inStreamReport);
                    inStreamReport.ReadText(Body);
                    //Body

                    EmailMessage.Create(Rec."Send To", 'Email Account', Body, true);

                    IF SHIPMENT.FindFirst() then;
                    clear(RecRef);

                    RecRef.GetTable(SHIPMENT);
                    FldRef := RecRef.Field(SHIPMENT.FieldNo("No."));
                    FldRef.SetRange(SHIPMENT."No.");

                    //Generate blob from report
                    TempBlob.CreateOutStream(outStreamReport);
                    TempBlob.CreateInStream(inStreamReport);
                    Report.SaveAs(Report::"Sales - Shipment", '', ReportFormat::Pdf, outStreamReport, RecRef);
                    AttachmentName := 'Shipment.pdf';
                    EmailMessage.AddAttachment(AttachmentName, 'Shipment', inStreamReport);

                    i := 1;


                    if SInvH.FindSet() then
                        repeat

                            TempBlobArray[i].CreateOutStream(ArrayOutReport[i]);
                            TempBlobArray[i].CreateInStream(ArrayInstrReport[i]);
                            FullFileName := 'Invoice' + SInvH."No." + '.' + 'pdf';
                            clear(RecRef);
                            RecRef.GetTable(SInvH);
                            FldRef := RecRef.Field(SInvH.FieldNo("No."));
                            FldRef.SetRange(SInvH."No.");
                            if RecRef.FindFirst() then begin

                                Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, ArrayOutReport[i], RecRef);

                                EmailMessage.AddAttachment(FullFileName, 'Invoice', ArrayInstrReport[i]);
                            end;
                            i += 1;

                        until (SInvH.Next = 0) or (i = 3);


                    if not Email.Send(EmailMessage, Enum::"Email Scenario"::Default) then begin
                        Message(GetLastErrorText);
                        ClearLastError();

                    end else
                        Message('Success');


                end;
            }

        }
    }

    local procedure GetImageFileExtension(var TenantMedia: Record "Tenant Media"): Text;
    begin
        case TenantMedia."Mime Type" of
            'image/jpeg':
                exit('.jpg');
            'image/bmp':
                exit('.bmp');
            'image/gif':
                exit('.gif');
            'image/png':
                exit('.png');
            'image/tiff':
                exit('.tiff');
            'application/msword':
                exit('.doc');
        //application/msword
        end;
    end;
    //https://docs.microsoft.com/en-us/dynamics365/business-central/dev-itpro/developer/methods-auto/media/media-importstream-instream-text-text-method
    //https://docs.microsoft.com/en-us/dynamics365/business-central/dev-itpro/developer/methods-auto/media/media-importfile-method
    //https://docs.microsoft.com/en-us/dynamics365/business-central/dev-itpro/developer/devenv-working-with-media-on-records#SupportedMediaTypes

    var

      
}
