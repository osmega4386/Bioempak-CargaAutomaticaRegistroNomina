tableextension 50510 "OSMG Gen. Journal Line" extends "Gen. Journal Line"
{
    fields
    {
        // Add changes to table fields here
        field(50100; "OSMG VAT Registration Type"; Code[10])
        {
            DataClassification = ToBeClassified;
        }
    }

    var
        myInt: Integer;
}