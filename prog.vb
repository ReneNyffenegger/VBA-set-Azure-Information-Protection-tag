option explicit

sub main(saveAsDir as string) ' {

    dim wb as workbook
    set wb = workbooks.add

    startHook
    wb.sheets(1).cells(1,1) = "Press ctrl-s to save"
    wb.sheets(1).cells(2,1) = "Then call stopHook()"

end sub ' }
