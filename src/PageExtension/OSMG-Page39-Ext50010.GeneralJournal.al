pageextension 50100 "OSMG General Journal" extends "General Journal"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        // Add changes to page actions here
        addlast(processing)
        {
            action("Export / Import Payroll")
            {
                ApplicationArea = Basic, Suite;
                Caption = 'Export / Import Payroll';
                Ellipsis = true;
                Image = Excel;
                Promoted = true;
                PromotedCategory = Process;

                trigger OnAction()
                var
                    ExportImportGenJournal: Report "Export / Import Gen Journal";
                begin
                    Clear(ExportImportGenJournal);
                    ExportImportGenJournal.SetParameters(Rec);
                    ExportImportGenJournal.Run();
                    CurrPage.Update();
                end;
            }
        }
    }

    var
        myInt: Integer;
}