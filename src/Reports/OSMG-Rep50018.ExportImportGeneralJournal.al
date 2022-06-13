report 50018 "Export / Import Gen Journal"
{
    UsageCategory = Administration;
    ApplicationArea = All;
    Caption = 'Export / Import Gen Journal';
    ProcessingOnly = true;

    dataset
    {
        dataitem(GenJournalLine; "Gen. Journal Line")
        {
            // column(ColumnName; SourceFieldName)
            // {

            // }

            trigger OnPreDataItem()
            begin
                if PrintToExcel then begin
                    MakeExcelDataHeader();
                    GenJnlLine.SetCurrentKey("Journal Template Name", "Journal Batch Name");
                    GenJnlLine.SetRange("Journal Template Name", JnlTmplateName);
                    GenJnlLine.SetRange("Journal Batch Name", JnlBatchName);
                    if GenJnlLine.Find('-') then
                        repeat
                            MakeExcelDataBody();
                        until GenJnlLine.NEXT = 0;
                    CurrReport.Break();
                    MakeExcel();
                end;
            end;

            trigger OnPostDataItem()
            begin
                if PrintToExcel then begin
                    MakeExcel;
                end;
            end;
        }
    }

    requestpage
    {
        layout
        {
            area(Content)
            {
                group(General)
                {
                    field(FileName; FileName)
                    {
                        ApplicationArea = All;
                        Caption = 'File Name';

                        trigger OnAssistEdit()
                        begin
                            UploadIntoStream(UploadExcelMsg, '', '', FromServerFileName, Istream);
                            if FromServerFileName <> '' then
                                FileName := FileManagement.GetFileName(FromServerFileName)
                            else
                                Error(FileErr);
                        end;
                    }
                    field(SheetName; SheetName)
                    {
                        ApplicationArea = All;
                        Caption = 'Sheet Name';

                        trigger OnAssistEdit()
                        begin
                            SheetName := TempExcelBuffer.SelectSheetsNameStream(Istream);
                            if SheetName = '' then
                                Error('');
                        end;
                    }
                    field(ShowPrint; ShowPrint)
                    {
                        ApplicationArea = All;
                        Caption = 'Operation Type';
                    }
                }
            }
        }

        // actions
        // {
        //     area(processing)
        //     {
        //         action(ActionName)
        //         {
        //             ApplicationArea = All;

        //         }
        //     }
        // }
    }

    // rendering
    // {
    //     layout(LayoutName)
    //     {
    //         Type = RDLC;
    //         LayoutFile = 'mylayout.rdl';
    //     }
    // }

    trigger OnPreReport()
    begin
        if JnlTmplateName = '' then
            Error(Text002);

        if JnlBatchName = '' then
            Error(Text003);

        TempExcelBuffer.LockTable();
        if ShowPrint = ShowPrint::"Import from Excel" then begin
            ReadExcelSheet();
            AnalyzeData();
        end;

        if ShowPrint = ShowPrint::"Print to Excel" then begin
            PrintToExcel := true;
        end;
    end;

    trigger OnPostReport()
    begin
        CLEAR(TempExcelBuffer);
    end;

    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        GenJnlLine: Record "Gen. Journal Line";
        GenJnlLine2: Record "Gen. Journal Line";
        Window: Dialog;
        FromServerFileName: Text;
        FileName: Text;
        Day: Text[30];
        Month: Text[30];
        Year: Text[30];
        SheetName: Text;
        ShowPrint: Enum "Operation Type";
        UploadExcelMsg: Label 'Select file..';
        Istream: InStream;
        FileManagement: Codeunit "File Management";
        FileErr: Label 'File does not exit';
        JnlTmplateName: Code[10];
        JnlBatchName: Code[10];
        PrintToExcel: Boolean;
        DataLbl: Label 'Data';
        ExcelFileName: Label '%1_%2_%3';
        Text001: Label 'General Journal';
        Text002: Label 'The Journal Template Name don''t be blank';
        Text003: Label 'The Journal Batch Name don''t be blank';
        Text006: Label 'Payment';
        Text007: Label 'Invoice';
        Text008: Label 'Credit Memo';
        Text009: Label 'Finance Charge Memo';
        Text010: Label 'Reminder';
        Text011: Label 'Refund';
        Text012: Label 'Debit Memo';
        Text013: Label 'G/L Account';
        Text014: Label 'Customer';
        Text015: Label 'Vendor';
        Text016: Label 'Bank Account';
        Text019: Label 'NIT';
        Text020: Label 'C.C.';
        Text021: Label 'C.C.+DV';
        Text022: Label 'C.E.';
        Text023: Label 'Foreign';
        Text024: Label 'Day';
        Text025: Label 'Month';
        Text026: Label 'Year';
        Text027: Label 'Amount (Integer Part)';
        Text028: Label 'Amount (Decimal Part)';
        Text029: Label 'Fixed Asset';
        Text030: Label 'IC Partner';
        RecNo: Integer;
        TotalRecNo: Integer;
        RecNo2: Integer;
        Band: Integer;
        LineNo: Integer;
        D: Integer;
        M: Integer;
        Y: Integer;

    /// <summary>
    /// SetParameters.
    /// </summary>
    /// <param name="GenJnlLine">Record "Gen. Journal Line".</param>
    procedure SetParameters(GenJnlLine: Record "Gen. Journal Line")
    begin
        JnlTmplateName := GenJnlLine."Journal Template Name";
        JnlBatchName := GenJnlLine."Journal Batch Name";
    end;

    local procedure ReadExcelSheet()
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        //TempExcelBuffer.OpenBook(FileName, SheetName);
        TempExcelBuffer.OpenBookStream(Istream, SheetName);
        TempExcelBuffer.ReadSheet;
    end;

    local procedure MakeExcelDataHeader()
    begin
        TempExcelBuffer.NewRow;
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Journal Template Name"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Journal Batch Name"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(FORMAT(Text024), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(FORMAT(Text025), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(FORMAT(Text026), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        //TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Posting Date"),false,'',true,false,false,'', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Document Type"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Document No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Account Type"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Account No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION(Description), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION(Amount), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Bal. Account Type"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Bal. Account No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("D1 VAT Registration Type"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("VAT Registration No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("D1 Bal. VAT Registration Type"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("D1 Bal. VAT Registration No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Shortcut Dimension 1 Code"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.FIELDCAPTION("Shortcut Dimension 2 Code"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure MakeExcelDataBody()
    begin
        TempExcelBuffer.NewRow;
        TempExcelBuffer.AddColumn(JnlTmplateName, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(JnlBatchName, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        Day := Format(GenJnlLine."Posting Date", 2, '<Day>');
        TempExcelBuffer.AddColumn(Day, false, '', false, false, false, '@', TempExcelBuffer."Cell Type"::Text);
        Month := Format(GenJnlLine."Posting Date", 0, '<Month>');
        TempExcelBuffer.AddColumn(Month, false, '', false, false, false, '@', TempExcelBuffer."Cell Type"::Text);
        Year := Format(GenJnlLine."Posting Date", 4, '<Year>');
        TempExcelBuffer.AddColumn(Year, false, '', false, false, false, '@', TempExcelBuffer."Cell Type"::Text);
        //TempExcelBuffer.AddColumn(GenJnlLine."Posting Date",false,'',false,false,false,'@');
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Document Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Account Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Account No."), false, '', false, false, false, '@', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine.Amount, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Bal. Account Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJnlLine."Bal. Account No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."VAT Registration No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."D1 VAT Registration Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."D1 Bal. VAT Registration No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."D1 Bal. VAT Registration Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Shortcut Dimension 1 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Format(GenJnlLine."Shortcut Dimension 2 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure MakeExcel()
    var
        myInt: Integer;
    begin
        TempExcelBuffer.CreateNewBook(Text001);
        TempExcelBuffer.WriteSheet(DataLbl + ' ' + Text001, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, Text001, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
    end;

    local procedure GetLine(): Integer
    var
        _GenJnlLine: Record "Gen. Journal Line";
    begin
        _GenJnlLine.Reset();
        _GenJnlLine.SetCurrentKey("Journal Template Name", "Journal Batch Name");
        _GenJnlLine.SetRange("Journal Template Name", JnlTmplateName);
        _GenJnlLine.SetRange("Journal Batch Name", JnlBatchName);
        if _GenJnlLine.FIND('+') then
            LineNo := _GenJnlLine."Line No." + 10000
        ELSE
            LineNo := 10000;

        EXIT(LineNo);
    end;

    local procedure AnalyzeData()
    begin
        Window.Open(
          Text001 +
          '@1@@@@@@@@@@@@@@@@@@@@@@@@@\');
        Window.Update(1, 0);
        TotalRecNo := TempExcelBuffer.COUNT;
        RecNo := 0;
        Band := 1;

        TempExcelBuffer.RESET;
        CLEAR(TempExcelBuffer);
        if TempExcelBuffer.FIND('-') then
            repeat
                RecNo := RecNo + 1;
                RecNo2 := TempExcelBuffer."Row No.";
                Window.UPDATE(1, ROUND(RecNo / TotalRecNo * 10000, 1));
                if (TempExcelBuffer."Row No." <> 1) then begin
                    if (TempExcelBuffer."Column No." = 1) then begin
                        GenJnlLine2.INIT;
                        GenJnlLine2."Journal Template Name" := JnlTmplateName;
                        GenJnlLine2."Journal Batch Name" := JnlBatchName;
                        GenJnlLine2."Line No." := GetLine;
                        GenJnlLine2.INSERT;
                    end;
                    InsertField(TempExcelBuffer."Column No.", TempExcelBuffer."Cell Value as Text");
                end;

            until TempExcelBuffer.NEXT = 0;
        Window.Close();
    end;

    local procedure InsertField(_Field: Integer; TextNoFormat: Text)
    var
        Amt: Decimal;
    begin
        with GenJnlLine2 do begin
            case _Field of
                3:
                    Evaluate(D, TextNoFormat);
                4:
                    Evaluate(M, TextNoFormat);

                5:
                    begin
                        Evaluate(Y, TextNoFormat);
                        Y += 2000;
                        "Posting Date" := DMY2Date(D, M, Y);
                    end;

                6:
                    begin
                        if TextNoFormat = Text006 then
                            "Document Type" := "Document Type"::Payment;
                        if TextNoFormat = Text007 then
                            "Document Type" := "Document Type"::Invoice;
                        if TextNoFormat = Text008 then
                            "Document Type" := "Document Type"::"Credit Memo";
                        if TextNoFormat = Text009 then
                            "Document Type" := "Document Type"::"Finance Charge Memo";
                        if TextNoFormat = Text010 then
                            "Document Type" := "Document Type"::Reminder;
                        if TextNoFormat = Text011 then
                            "Document Type" := "Document Type"::Refund;
                        if TextNoFormat = Text012 then
                            "Document Type" := "Document Type"::"Debit Memo";
                    end;

                7:
                    "Document No." := DelChr(TextNoFormat, '>', ' ');
                8:
                    begin
                        if TextNoFormat = Text013 then
                            "Account Type" := "Account Type"::"G/L Account";
                        if TextNoFormat = Text014 then
                            "Account Type" := "Account Type"::Customer;
                        if TextNoFormat = Text015 then
                            "Account Type" := "Account Type"::Vendor;
                        if TextNoFormat = Text016 then
                            "Account Type" := "Account Type"::"Bank Account";
                        if TextNoFormat = Text029 then
                            "Account Type" := "Account Type"::"Fixed Asset";
                        if TextNoFormat = Text030 then
                            "Account Type" := "Account Type"::"IC Partner";
                    end;
                9:
                    Validate("Account No.", DelChr(TextNoFormat, '=', '.,'));
                10:
                    Description := DelChr(TextNoFormat, '>', ' ');
                11:
                    Validate("Posting Group", DelChr(TextNoFormat, '>', ' '));

                12:
                    begin
                        Evaluate(Amt, DelChr(TextNoFormat, '>', ' '));
                        Validate(Amount, Round(Amt));
                    end;
                13:
                    begin
                        if TextNoFormat = Text013 then
                            "Bal. Account Type" := "Bal. Account Type"::"G/L Account";
                        if TextNoFormat = Text014 then
                            "Bal. Account Type" := "Bal. Account Type"::Customer;
                        if TextNoFormat = Text015 then
                            "Bal. Account Type" := "Bal. Account Type"::Vendor;
                        if TextNoFormat = Text016 then
                            "Bal. Account Type" := "Bal. Account Type"::"Bank Account";
                        if TextNoFormat = Text029 then
                            "Bal. Account Type" := "Bal. Account Type"::"Fixed Asset";
                        if TextNoFormat = Text030 then
                            "Bal. Account Type" := "Bal. Account Type"::"IC Partner";
                    end;
                14:
                    "Bal. Account No." := DelChr(TextNoFormat, '>', ' ');
                15:
                    "VAT Registration No." := DelChr(TextNoFormat, '>', ' ');

                16:
                    begin

                        // if TextNoFormat = Text019 then
                        //     "VAT Registration Type" := "VAT Registration Type"::NIT;
                        // if TextNoFormat = Text020 then
                        //     "VAT Registration Type" := "VAT Registration Type"::"C.C.";
                        // if TextNoFormat = Text021 then
                        //     "VAT Registration Type" := "VAT Registration Type"::"C.C.+DV";
                        // if TextNoFormat = Text022 then
                        //     "VAT Registration Type" := "VAT Registration Type"::"C.E.";
                        // if TextNoFormat = Text023 then
                        //     "VAT Registration Type" := "VAT Registration Type"::Foreign;

                        Validate("D1 VAT Registration Type", TextNoFormat);
                    end;

                17:
                    "D1 Bal. VAT Registration No." := DELCHR(TextNoFormat, '>', ' ');

                18:

                    begin
                        //   if TextNoFormat = Text019 then
                        //                     "Bal. VAT Registration Type" := "Bal. VAT Registration Type"::asd;
                        //                 if TextNoFormat = Text020 then
                        //                     "Bal. VAT Registration Type" := "Bal. VAT Registration Type"::"2";
                        //                 if TextNoFormat = Text021 then
                        //                     "Bal. VAT Registration Type" := "Bal. VAT Registration Type"::"3";
                        //                 if TextNoFormat = Text022 then
                        //                     "Bal. VAT Registration Type" := "Bal. VAT Registration Type"::"4";
                        //                 if TextNoFormat = Text023 then
                        //                     "Bal. VAT Registration Type" := "Bal. VAT Registration Type"::"5";

                        Validate("D1 Bal. VAT Registration Type", TextNoFormat);
                    end;

                19:
                    Validate("Shortcut Dimension 1 Code", TextNoFormat);
                20:
                    Validate("Shortcut Dimension 2 Code", TextNoFormat);
            end;
            Modify();
        end;
    end;

    local procedure GetDate(StrText: Text): Date
    var
        Day: Integer;
        Month: Integer;
        Year: Integer;
    begin
        Evaluate(Day, CopyStr(StrText, 1, 2));
        Evaluate(Month, CopyStr(StrText, 4, 2));
        Evaluate(Year, CopyStr(StrText, 7, 2));
        Year := Year + 2000;
        exit(DMY2Date(Day, Month, Year));
    end;
}