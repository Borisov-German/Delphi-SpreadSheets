// ********************************************
// ** Object for dual  SpreadSheet  managing **
// ** using Excel or OpenOffice automaticaly **
// ** By: Sergio Hernandez                   **
// ** oficina(at)hcsoft.net, CopyLeft 2004   **
// ** Version 0.92 18-05-2004 (DDMMYYYY)     **
// ** Use it freely, change it, etc.         **
// ********************************************

// ********************************************
// ** Modify By: Romulo Silva Ramos          **
// ** Version 0.93 01-10-2009 (DDMMYYYY)     **
// ********************************************
// http://www.oooforum.org/forum/viewtopic.phtml?t=8878

// ********************************************
// ** Modify By: Borisov German              **
// ** Version 0.94 25-01-2011 (DDMMYYYY)     **
// ********************************************

{EXAMPLE OF USE
  //Create object: We have two flavours:
  //(A) from an existing file...
  HCalc:= TSpreadSheet.create(OpenDialog.FileName, false);
  //(B) from a blank document...
  HCalc:= TSpreadSheet.create(sstOpenOffice); //OpenOffice doc if possible, please
  HCalc.FileName:= 'C:\MyNewDoc'; //Needs a file name before you SaveDoc!
  //--end of creation.
  HCalc.ActivateSheetByIndex(2); //Activate second sheet
  if HCalc.GetActiveSheetProtected then
    ShowMessage('2nd sheet of name "'+HCalc.ActiveSheetName+'" IS protected');
  //Change a cell value (well, change formula, not the double float value)
  if HCalc.CellText[i,2] = '' then HCalc.CellText[i,2] := 'Hello world!';
  HCalc.AddNewSheet('New Sheet');
  HCalc.PrintDoc;
  HCalc.SaveDoc;
  HCalc.Free;

  Other exemple:
       var
          Plan : TSpreadSheet;
       begin
          Plan := TSpreadSheet.Create(sstExcel,false);
          Plan.FileName := arquivo;

          Plan.ActivateSheetByIndex(1);

          Plan.CellText[1,1] := 'Abstraction complete';

          Plan.SaveDoc;
          Plan.Visible := True;
       begin
}

{TODO LIST:
  -PrintActiveSheet is not working for OpenOffice (is it even possible?)
  -A way to write a date in a cell changing also the format (Excel is herratic in that)
}

{CHANGE LOG:
 V0.94
   -Many renames

 V0.93:
   -SetFontSize(Row, Col, Size): change font size in that cell.
   -Change AddNewSheet to add a new sheet in end at sheet list
   -Add ValidateSheetName to validate sheet names when adding or renaming a sheet
   -Change in SaveDoc to use SaveAs/StoreAsUrl
   -Change name 'Hoja' para 'Sheet'
   -Change name 'Programa' para 'SheetSoftware'
   -Change exception.create messages:
       'No puedo leer el fichero "'+FileName+'" al no estar presente el programa necesario.'
       to
       'Not able to read the file "'+FileName+'". Probably not exist a software installed to open it.'
   -Add other exemple of utilization
 V0.92:
   -SetActiveSheetName didn't change the name to the right sheet on OpenOffice.
   -PrintPreview: New procedure to show up the print preview window.
   -SetBold(Row, Col): Make SetCellBold the text in that cell.
   -ColumnWidth(col, width): To change a column width.
 V0.91:
   -NewDoc: New procedure for creating a blank doc (used in create)
   -Create from empty doc adds a blank document and take visibility as parameter.
   -New functions ooCreateValue and ooDispatch to clean up the code.
   -ActiveSheetName: Now is a read-write property, not a read-only function.
   -Visible: Now is a read-write property instead of a create param only.
 V0.9:
  -Create from empty doc now tries both programs (if OO fails try to use Excel).
  -CellTextByName: Didn't work on Excel docs.
}

Unit uSpreadsheets;

Interface

Uses Variants, SysUtils, ComObj, Graphics, ActiveX;

//sstError: Tried to open but both failes
//sstNone:  Haven't tried still to open any
Type
  TSpreadSheetType = (sstError, sstNone, sstExcel, sstOpenOffice);

  TCellHorizontalAlign = (chaDefault, chaLeft, chaCenter, chaRight, chaBoth);

  TCellVerticalAlign = (cvaDefault, cvaTop, cvaCenter, cvaBottom);

  TCellBorder = (cbLeft, cbRight, cbTop, cbBottom);

  TCellBorders = Set Of TCellBorder;

  { TODO : Add other formats, that supported by both programm}
  TFileFormat = (ffExcel97);

Const
  { TODO : replace hardcoded values to OLE constants }
  // OpenOffice.org_3.0_SDK\sdk\idl\com\sun\star\table\CellHoriJustify.idl

  ooCellHoriJustify: Array[TCellHorizontalAlign] Of TOleEnum = (0, 1, 2, 3, 4);
  ooCellVertJustify: Array[TCellVerticalAlign] Of TOleEnum = (0, 1, 2, 3);

  XlHAlign: Array[TCellHorizontalAlign] Of TOleEnum = ($00000001, $FFFFEFDD, $FFFFEFF4, $FFFFEFC8, $00000005);
  XlVAlign: Array[TCellVerticalAlign] Of TOleEnum = (0, $FFFFEFC0, $FFFFEFF4, $FFFFEFF5);

  // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlfileformat.aspx
  XlFileFormat: Array[TFileFormat] Of TOleEnum = (56);
  ooFileFormat: Array[TFileFormat] Of String = ('MS Excel 97');

  cbsNone: TCellBorders = [];
  cbsAll: TCellBorders = [cbLeft, cbRight, cbTop, cbBottom];

  // Excel mesure heightness in typographic point (1/72 inch), OOo hundredth of mm.
  OOoToXlHeight: Double = 72 / 25.4 / 100;
Var
  //Excel mesure width in count of overage char width in defaukt font
  // this coefficient is variable and must be recalculated when default font is chanded
  OOoToXlWidth: Double;

Type
  TSpreadSheet = Class(TObject)
  Private
    FVisible: Boolean;
    //Program loaded stuff...
    Procedure LoadProg;
    Procedure CloseProg;
    Function GetProgLoaded: Boolean;
    Function GetDocLoaded: Boolean;
    Function GetIsExcel: Boolean;
    Function GetIsOpenOffice: Boolean;
    Procedure SetVisible(Value: Boolean);
    //Sheets stuff..
    Function GetCountSheets: Integer;
    Function GetActiveSheetName: String;
    Procedure SetActiveSheetName(NewName: String);
    //Cells stuff...
    Function GetCellText(Row, Col: Integer): String;
    Procedure SetCellText(Row, Col: Integer; Value: String);
    Function GetCellTextByName(Range: String): String;
    Procedure SetCellTextByName(Range: String; Value: String);
    //OpenOffice only stuff...
    Function FileName2URL(FileName: String): String;
    Procedure ooDispatch(ooCommand: String; ooParams: Variant);
    Function ooCreateValue(ooName: String; ooData: Variant): Variant;
    // Excel only stuff...
    Function xlRange(Row1, Col1, Row2, Col2: Integer): Variant;
    Function xlGetWidthMultiplier: Double;
    //aux functions
    Function ValidateSheetName(Name: String): String;
  Public
    Tipo: TSpreadSheetType; //Witch program was used to manage the doc?
    FileName: String; //In windows FileName format C:\MyDoc.XXX
    SheetSoftware: Variant; //Excel or OpenOfice instance created.
    Desktop: Variant; //OpenOffice desktop reference (not used now).
    Document: Variant; //Document opened.
    ActiveSheet: Variant; //Active sheet.
    // TODO move to Private section
    Function ooCreateUnoStruct(StructName: String; IndexMax: Integer = -1): Variant;
    //Object internals...
    Constructor Create(Name: String; MakeVisible: Boolean); Overload;
    Constructor Create(MyTipo: TSpreadSheetType; MakeVisible: Boolean); Overload;
    Destructor Destroy; Override;
    //Program loaded stuff...
    Procedure NewDoc;
    Procedure LoadDoc;
    Procedure CloseDoc;

    Function SaveDoc(FileFormat: TFileFormat = ffExcel97): Boolean;
    Function PrintDoc: Boolean;
    Procedure ShowPrintPreview;
    Property ProgLoaded: Boolean Read GetProgLoaded;
    Property DocLoaded: Boolean Read GetDocLoaded;
    Property IsExcel: Boolean Read GetIsExcel;
    Property IsOpenOffice: Boolean Read GetIsOpenOffice;
    Property Visible: Boolean Read FVisible Write SetVisible;
    //Sheets stuff...
    Function ActivateSheetByIndex(nIndex: Integer): Boolean;
    Function ActivateSheetByName(SheetName: String; CaseSensitive: Boolean): Boolean;
    Function GetActiveSheetProtected: Boolean;
    Function ChangeActiveSheetProtection(IsProtected: Boolean; Password: String): Boolean;
    Function PrintActiveSheet: Boolean;
    Procedure AddNewSheet(NewName: String);
    Property CountSheets: Integer Read GetCountSheets;
    Property ActiveSheetName: String Read GetActiveSheetName Write SetActiveSheetName;
    Procedure SetActiveSheetFont(Font: String);
    Procedure SetActiveSheetFontSize(Size: Integer);
    // Region stuff...
    Procedure SetRegionMerge(Row1, Col1, Row2, Col2: Integer; Value: Boolean);
    Procedure SetRegionWordWrap(Row1, Col1, Row2, Col2: Integer; Value: Boolean);
    Procedure SetRegionBorder(Row1, Col1, Row2, Col2: Integer; Value: TCellBorders);
    Procedure SetRegionProtected(Row1, Col1, Row2, Col2: Integer; Value: Boolean);
    //Cells stuff...
    Procedure SetCellBackgroundColor(Row, Col: Integer; Color: TColor);
    Procedure SetCellFont(Row, Col: Integer; Font: String);
    Procedure SetCellFontSize(Row, Col, Size: Integer);
    Procedure SetCellBold(Row, Col: Integer; Value: Boolean);
    Procedure SetCellVerticalAlign(Row, Col: Integer; Align: TCellVerticalAlign);
    Procedure SetCellHorizontalAlign(Row, Col: Integer; Align: TCellHorizontalAlign);
    Procedure SetCellWordWrap(Row, Col: Integer; Value: Boolean);
    Procedure SetCellFloat(Row, Col: Integer; Value: Double);
    Procedure SetCellProtected(Row, Col: Integer; Value: Boolean);

    Property CellText[Row, Col: Integer]: String Read GetCellText Write SetCellText;
    Property CellTextByName[Range: String]: String Read GetCellTextByName Write SetCellTextByName;
    //Other stuff
    Procedure SetColumnWidth(Col, Width: Integer); //Width in 1/100 of mm.
    Procedure SetRowHeight(Row, Height: Integer); //Height in 1/100 of mm.
  End;

Implementation

Uses Windows;

// ************************
// ** Create and destroy **
// ************************

//Create with an empty doc of requested type (use sstExcel or sstOpenOffice)
//Remember to define FileName before calling to SaveDoc

Constructor TSpreadSheet.Create(MyTipo: TSpreadSheetType; MakeVisible: Boolean);
Var
  I: Integer;
  IsFirstTry: Boolean;
Begin
  //Close all opened things first...
  If DocLoaded Then CloseDoc;
  If ProgLoaded Then CloseProg;
  //I will try to open twice, so if Excel fails, OpenOffice is used instead
  IsFirstTry := True;
  For I := 1 To 2 Do
    Begin
      //Try to open OpenOffice...
      If (MyTipo = sstOpenOffice) Or (MyTipo = sstNone) Then
        Begin
          SheetSoftware := CreateOleObject('com.sun.star.ServiceManager');
          If ProgLoaded Then
            Begin
              Tipo := sstOpenOffice;
              break;
            End
          Else
            Begin
              If IsFirstTry Then
                Begin
                  //Try Excel as my second choice
                  MyTipo := sstExcel;
                  IsFirstTry := False;
                End
              Else
                Begin
                  //Both failed!
                  break;
                End;
            End;
        End;
      //Try to open Excel...
      If (MyTipo = sstExcel) Or (MyTipo = sstNone) Then
        Begin
          SheetSoftware := CreateOleObject('Excel.Application');
          If ProgLoaded Then
            Begin
              Tipo := sstExcel;
              break;
            End
          Else
            Begin
              If IsFirstTry Then
                Begin
                  //Try OpenOffice as my second choice
                  MyTipo := sstOpenOffice;
                  IsFirstTry := False;
                End
              Else
                Begin
                  //Both failed!
                  break;
                End;
            End;
        End;
    End;
  //Was it able to open any of them?
  If Tipo = sstNone Then
    Begin
      Tipo := sstError;
      Raise Exception.Create('TSheetCalc.create failed, may be no Office is installed?');
    End;
  //Add a blank document...
  FVisible := MakeVisible;
  NewDoc;
End;

Constructor TSpreadSheet.Create(Name: String; MakeVisible: Boolean);
Begin
  //Store values...
  FileName := Name;
  FVisible := MakeVisible;
  //Open program and document...
  LoadProg;
  LoadDoc;
End;

Destructor TSpreadSheet.Destroy;
Begin
  CloseDoc;
  CloseProg;
  Inherited;
End;

// *************************
// ** Loading the program **
// ** Excel or OpenOffice **
// *************************

Procedure TSpreadSheet.LoadProg;
Begin
  If ProgLoaded Then CloseProg;
  If (UpperCase(ExtractFileExt(FileName)) = '.XLS') Then
    Begin
      //Excel is the primary choice...
      SheetSoftware := CreateOleObject('Excel.Application');
      If ProgLoaded Then Tipo := sstExcel;
    End;
  //Not lucky with Excel? Another filetype? Let's go with OpenOffice...
  If Tipo = sstNone Then
    Begin
      //Try with OpenOffice...
      SheetSoftware := CreateOleObject('com.sun.star.ServiceManager');
      If ProgLoaded Then Tipo := sstOpenOffice;
    End;
  //Still no program loaded?
  If Not ProgLoaded Then
    Begin
      Tipo := sstError;
      Raise Exception.Create('TSheetCalc.create failed, may be no Office is installed?');
    End;
End;

Procedure TSpreadSheet.CloseProg;
Begin
  If DocLoaded Then CloseDoc;
  If ProgLoaded Then
    Begin
      Try
        If IsExcel Then SheetSoftware.Quit;
        SheetSoftware := Unassigned;
      Finally
      End;
    End;
  Tipo := sstNone;
End;

//Is there any prog loaded? Witch one?

Function TSpreadSheet.GetProgLoaded: Boolean;
Begin
  Result := Not (VarIsEmpty(SheetSoftware) Or VarIsNull(SheetSoftware));
End;

Function TSpreadSheet.GetIsExcel: Boolean;
Begin
  Result := (Tipo = sstExcel);
End;

Function TSpreadSheet.GetIsOpenOffice: Boolean;
Begin
  Result := (Tipo = sstOpenOffice);
End;

// ************************
// ** Loading a document **
// ************************

Procedure TSpreadSheet.NewDoc;
Var
  ooParams: Variant;
Begin
  //Is the program running? (Excel or OpenOffice)
  If Not ProgLoaded Then Raise Exception.Create('No program loaded for the new document.');
  //Is there a doc already loaded?
  If DocLoaded Then CloseDoc;
  Desktop := Unassigned;
  //OK, now try to create the doc...
  If IsExcel Then
    Begin
      SheetSoftware.WorkBooks.Add;
      SheetSoftware.Visible := Visible;
      SheetSoftware.DisplayAlerts := False;
      Document := SheetSoftware.ActiveWorkBook;
      ActiveSheet := Document.ActiveSheet;
      OOoToXlWidth := xlGetWidthMultiplier;
    End
  Else If IsOpenOffice Then
    Begin
      Desktop := SheetSoftware.CreateInstance('com.sun.star.frame.Desktop');
      //Optional parameters (visible)...
      ooParams := VarArrayCreate([0, 0], varVariant);
      ooParams[0] := ooCreateValue('Hidden', Not Visible);
      //Create the document...
      Document := Desktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, ooParams);
      ActivateSheetByIndex(1);
    End;
End;

Procedure TSpreadSheet.LoadDoc;
Var
  ooParams: Variant;
Begin
  If FileName = '' Then Exit;
  //Is the program running? (Excel or OpenOffice)
  If Not ProgLoaded Then LoadProg;
  //Is there a doc already loaded?
  If DocLoaded Then CloseDoc;
  Desktop := Unassigned;
  //OK, now try to open the doc...
  If IsExcel Then
    Begin
      SheetSoftware.WorkBooks.Open(FileName, 3);
      SheetSoftware.Visible := Visible;
      SheetSoftware.DisplayAlerts := False;
      Document := SheetSoftware.ActiveWorkBook;
      ActiveSheet := Document.ActiveSheet;
      OOoToXlWidth := xlGetWidthMultiplier;
    End
  Else If IsOpenOffice Then
    Begin
      Desktop := SheetSoftware.CreateInstance('com.sun.star.frame.Desktop');
      //Optional parameters (visible)...
      ooParams := VarArrayCreate([0, 0], varVariant);
      ooParams[0] := ooCreateValue('Hidden', Not Visible);
      //Open the document...
      Document := Desktop.LoadComponentFromURL(FileName2URL(FileName), '_blank', 0, ooParams);
      ActivateSheetByIndex(1);
    End;
  If Tipo = sstNone Then
    Raise Exception.Create('Not able to read the file "' + FileName + '". Probably not exist a software installed to open it.');
End;

Function TSpreadSheet.SaveDoc(FileFormat: TFileFormat = ffExcel97): Boolean;
Var
  ooParams: Variant;
  Dir: String;
Begin
  Result := False;
  Dir := ExtractFilePath(FileName);
  If Not DirectoryExists(Dir) Then ForceDirectories(Dir);

  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Document.SaveAs(FileName, XlFileFormat[FileFormat]);
          Result := True;
        End
      Else If IsOpenOffice Then
        Begin
          ooParams := VarArrayCreate([0, 0], varVariant);
          ooParams[0] := ooCreateValue('FilterName', ooFileFormat[FileFormat]);
          Document.StoreAsUrl(FileName2URL(FileName), ooParams);
          Result := True;
        End;
    End;
End;

//Print the Doc...

Function TSpreadSheet.PrintDoc: Boolean;
Var
  ooParams: Variant;
Begin
  Result := False;
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Document.PrintOut;
          Result := True;
        End
      Else If IsOpenOffice Then
        Begin
          //NOTE: OpenOffice will print all sheets with Printable areas, but if no
          //printable areas are defined in the doc, it will print all entire sheets.
          //Optional parameters (wait until fully sent to printer)...
          ooParams := VarArrayCreate([0, 0], varVariant);
          ooParams[0] := ooCreateValue('Wait', True);
          Document.Print(ooParams);
          Result := True;
        End;
    End;
End;

Procedure TSpreadSheet.ShowPrintPreview;
Begin
  If DocLoaded Then
    Begin
      //Force visibility of the doc...
      Visible := True;
      If IsExcel Then
        Begin
          // TODO replase first 3 parameters by default value instead of null
          Document.PrintOut(Null, Null, Null, True);
        End
      Else If IsOpenOffice Then
        Begin
          ooDispatch('.uno:PrintPreview', Unassigned);
        End;
    End;
End;

Procedure TSpreadSheet.SetVisible(Value: Boolean);
Begin
  If DocLoaded And (Value <> FVisible) Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.Visible := Value;

        End
      Else If IsOpenOffice Then
        Begin
          Document.getCurrentController.getFrame.getContainerWindow.SetVisible(Value);
        End;
      FVisible := Value;
    End;
End;

Procedure TSpreadSheet.CloseDoc;
Begin
  If DocLoaded Then
    Begin
      //Close it...
      Try
        If IsOpenOffice Then
          Begin
            Document.Dispose;

          End
        Else If IsExcel Then
          Begin
            Document.Close;
          End;
      Finally
      End;
      //Clean up both "pointer"...
      Document := Null;
      ActiveSheet := Null;
    End;
End;

Function TSpreadSheet.GetDocLoaded: Boolean;
Begin
  Result := Not (VarIsEmpty(Document) Or VarIsNull(Document));
End;

// *********************
// ** Managing sheets **
// *********************

Function TSpreadSheet.GetCountSheets: Integer;
Begin
  Result := 0;
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Result := Document.Sheets.Count;
        End
      Else If IsOpenOffice Then
        Begin
          Result := Document.getSheets.GetCount;
        End;
    End;
End;

//Index is 1 based in Excel, but OpenOffice uses it 0-based
//Here we asume 1-based so OO needs to activate (nIndex-1)

Function TSpreadSheet.ActivateSheetByIndex(nIndex: Integer): Boolean;
Begin
  Result := False;
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Document.Sheets[nIndex].activate;
          ActiveSheet := Document.ActiveSheet;
          Result := True;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet := Document.getSheets.getByIndex(nIndex - 1);
          Result := True;
        End;
      sleep(100); //Asyncronus, so better give it time to make the change
    End;
End;

//Find a sheet by its name...

Function TSpreadSheet.ActivateSheetByName(SheetName: String; CaseSensitive: Boolean): Boolean;
Var
  OldActiveSheet: Variant;
  I: Integer;
Begin
  Result := False;
  If DocLoaded Then
    Begin
      If CaseSensitive Then
        Begin
          //Find the EXACT name...
          If IsExcel Then
            Begin
              Document.Sheets[SheetName].Select;
              ActiveSheet := Document.ActiveSheet;
              Result := True;
            End
          Else If IsOpenOffice Then
            Begin
              ActiveSheet := Document.getSheets.getByName(SheetName);
              Result := True;
            End;
        End
      Else
        Begin
          //Find the Sheet regardless of the case...
          OldActiveSheet := ActiveSheet;
          For I := 1 To GetCountSheets Do
            Begin
              ActivateSheetByIndex(I);
              If UpperCase(ActiveSheetName) = UpperCase(SheetName) Then
                Begin
                  Result := True;
                  Exit;
                End;
            End;
          //If not found, let the old active sheet active...
          ActiveSheet := OldActiveSheet;
        End;
    End;
End;

//Name of the active sheet?

Function TSpreadSheet.GetActiveSheetName: String;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Result := ActiveSheet.Name;
        End
      Else If IsOpenOffice Then
        Begin
          Result := ActiveSheet.GetName;
        End;
    End;
End;

Procedure TSpreadSheet.SetActiveSheetFont(Font: String);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells.Font.Name := Font;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.CharFontName := Font;
        End;
    End;
End;

Procedure TSpreadSheet.SetActiveSheetFontSize(Size: Integer);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells.Font.Size := Size;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.CharHeight := Size;
        End;
    End;
End;

Procedure TSpreadSheet.SetActiveSheetName(NewName: String);
Begin
  NewName := ValidateSheetName(NewName);

  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Name := NewName;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.setName(NewName);
          //This code always changes the name of "visible" sheet, not active one!
          //ooParams:= VarArrayCreate([0, 0], varVariant);
          //ooParams[0]:= ooCreateValue('Name', NewName);
          //ooDispatch('.uno:RenameTable', ooParams);
        End;
    End;
End;

Function TSpreadSheet.ChangeActiveSheetProtection(IsProtected: Boolean;
  Password: String): Boolean;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          If IsProtected Then
            ActiveSheet._Protect(Password)
          Else
            ActiveSheet.Unprotect(Password);
        End
      Else If IsOpenOffice Then
        Begin
          If IsProtected Then
            ActiveSheet.Protect(Password)
          Else
            ActiveSheet.Unprotect(Password);
        End;
    End;
  Result := GetActiveSheetProtected;
End;

//Check for sheet protection (password)...

Function TSpreadSheet.GetActiveSheetProtected: Boolean;
Begin
  Result := False;
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Result := ActiveSheet.ProtectContents;
        End
      Else If IsOpenOffice Then
        Begin
          Result := ActiveSheet.IsProtected;
        End;
    End;
End;

//WARNING: This function is NOT dual, only works for Excel docs!
//Send active sheet to default printer (as seen in preview window)...

Function TSpreadSheet.PrintActiveSheet: Boolean;
Begin
  Result := False;
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          ActiveSheet.PrintOut;
          Result := True;
        End
      Else If IsOpenOffice Then
        Begin
          Raise Exception.Create('Function "PrintActiveSheet" still not working in OpenOffice!');
          //ActiveSheet.Print;
          Result := False;
        End;
    End;
End;

//Add a new sheet, name it, and make it the active sheet...

Procedure TSpreadSheet.AddNewSheet(NewName: String);
Var
  ooSheets: Variant;
Begin
  NewName := ValidateSheetName(NewName);

  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Document.WorkSheets.Add;
          Document.ActiveSheet.Move(After := Document.Sheets[CountSheets]);
          Document.ActiveSheet.Name := NewName;
          //Active sheet has move to this new one, so I need to update the var
          ActiveSheet := Document.ActiveSheet;
        End
      Else If IsOpenOffice Then
        Begin
          ooSheets := Document.getSheets;
          ooSheets.insertNewByName(NewName, CountSheets);
          //Redefine active sheet to this new one
          ActiveSheet := ooSheets.getByName(NewName);
        End;
    End;
End;

// ************************
// ** Manage  the  cells **
// ** in the ActiveSheet **
// ************************

//Read/Write cell text (formula en Excel) by index
//OpenOffice start at cell (0,0) while Excel at (1,1)
//Also, Excel uses (row, col) and OpenOffice uses (col, row)

Function TSpreadSheet.GetCellText(Row, Col: Integer): String;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Result := ActiveSheet.Cells[Row, Col].Formula; //.Text;
        End
      Else If IsOpenOffice Then
        Begin
          Result := ActiveSheet.getCellByPosition(Col - 1, Row - 1).getFormula;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellText(Row, Col: Integer; Value: String);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          ActiveSheet.Cells[Row, Col].Formula := Value;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).setFormula(Value);
        End;
    End;
End;

//Read/Write cell text (formula in excel) by name instead of position
//For instance, you can set the value for cell 'NewSheet!A12' or similar
//NOTE: If range contains several cells, first one will be used.

Function TSpreadSheet.GetCellTextByName(Range: String): String;
Var
  OldActiveSheet: Variant;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Result := SheetSoftware.Range[Range].Text; //Set 'Formula' but Get 'Text';
        End
      Else If IsOpenOffice Then
        Begin
          OldActiveSheet := ActiveSheet;
          //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
          If pos('!', Range) > 0 Then
            Begin
              //Activate the proper sheet...
              If Not ActivateSheetByName(Copy(Range, 1, pos('!', Range) - 1), False) Then
                Raise Exception.Create('Sheet "' + Copy(Range, 1, pos('!', Range) - 1) + '" not present in the document.');
              Range := Copy(Range, pos('!', Range) + 1, 999);
            End;
          Result := ActiveSheet.getCellRangeByName(Range).getCellByPosition(0, 0).getFormula;
          ActiveSheet := OldActiveSheet;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellTextByName(Range: String; Value: String);
Var
  OldActiveSheet: Variant;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.Range[Range].Formula := Value;
        End
      Else If IsOpenOffice Then
        Begin
          OldActiveSheet := ActiveSheet;
          //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
          If pos('!', Range) > 0 Then
            Begin
              //Activate the proper sheet...
              If Not ActivateSheetByName(Copy(Range, 1, pos('!', Range) - 1), False) Then
                Raise Exception.Create('Sheet "' + Copy(Range, 1, pos('!', Range) - 1) + '" not present in the document.');
              Range := Copy(Range, pos('!', Range) + 1, 999);
            End;
          ActiveSheet.getCellRangeByName(Range).getCellByPosition(0, 0).setFormula(Value);
          ActiveSheet := OldActiveSheet;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellFloat(Row, Col: Integer; Value: Double);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          ActiveSheet.Cells[Row, Col].Formula := Value;
        End
      Else If IsOpenOffice Then
        Begin
          If IsOpenOffice Then ActiveSheet.getCellByPosition(Col - 1, Row - 1).setValue(Value);
        End;
    End;
End;

Procedure TSpreadSheet.SetCellVerticalAlign(Row, Col: Integer;
  Align: TCellVerticalAlign);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].VerticalAlignment := XlVAlign[Align];
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).VertJustify := ooCellVertJustify[Align];
        End;
    End;
End;

Procedure TSpreadSheet.SetCellWordWrap(Row, Col: Integer; Value: Boolean);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].WrapText := Value;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).IsTextWrapped := Value;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellBackgroundColor(Row, Col: Integer; Color: TColor);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].Interior.Color := Color;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).CellBackColor := Color;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellFont(Row, Col: Integer; Font: String);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].Font.Name := Font;
        End
      Else If
        IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).CharFontName := Font;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellFontSize(Row, Col, Size: Integer);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].Font.Size := Size;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).CharHeight := Size;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellHorizontalAlign(Row, Col: Integer;
  Align: TCellHorizontalAlign);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].HorizontalAlignment := XlHAlign[Align];
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).HoriJustify := ooCellHoriJustify[Align];
        End;
    End;
End;

Procedure TSpreadSheet.SetCellProtected(Row, Col: Integer; Value: Boolean);
Var
  CellProtection: Variant;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          If IsExcel Then ActiveSheet.Cells[Row, Col].Locked := Value;
        End
      Else If IsOpenOffice Then
        Begin
          CellProtection := ooCreateUnoStruct('com.sun.star.util.CellProtection');
          CellProtection.IsLocked := Value;
          ActiveSheet.getCellByPosition(Col - 1, Row - 1).CellProtection := CellProtection;
        End;
    End;
End;

Procedure TSpreadSheet.SetCellBold(Row, Col: Integer; Value: Boolean);
Const
  //OpenOffice.org_3.0_SDK\sdk\idl\com\sun\star\awt\FontWeight.idl
  ooBold: Integer = 150; //com.sun.star.awt.FontWeight.BOLD
  ooNorman: Integer = 100; //com.sun.star.awt.FontWeight.NORMAL
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Cells[Row, Col].Font.Bold := Value;
        End
      Else If IsOpenOffice Then
        Begin
          If Value Then
            ActiveSheet.getCellByPosition(Col - 1, Row - 1).CharWeight := ooBold
          Else
            ActiveSheet.getCellByPosition(Col - 1, Row - 1).CharWeight := ooNorman;
        End;
    End;
End;

Procedure TSpreadSheet.SetColumnWidth(Col, Width: Integer); //Width in 1/100 of mm.
Var
  Column: Variant;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          //Excel use the width of '0' as the unit, we do an aproximation: Width '0' = 2 mm.
          { TODO : find how to get exactly value of this const}

          SheetSoftware.ActiveSheet.Columns[Col].ColumnWidth := Width * OOoToXlWidth;
        End
      Else If IsOpenOffice Then
        Begin
          Column := ActiveSheet.getCellByPosition(Col - 1, 0).getColumns.getByIndex(0);
          Column.Width := Width;
          Column.IsVisible := (Width > 0);
        End;
    End;
End;

Procedure TSpreadSheet.SetRegionBorder(Row1, Col1, Row2, Col2: Integer;
  Value: TCellBorders);
Var
  Range, BorderOn, BOrderOff, Border: Variant;
Const
  xlThin = $00000002;
  xlContinuous = $00000001;
  xlInsideHorizontal = $0000000C;
  xlInsideVertical = $0000000B;
  xlDiagonalDown = $00000005;
  xlDiagonalUp = $00000006;
  xlEdgeBottom = $00000009;
  xlEdgeLeft = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop = $00000008;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          Range := xlRange(Row1, Col1, Row2, Col2);

          If cbLeft In Value Then
            Begin
              Border := Range.Borders[xlEdgeLeft];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlEdgeLeft].Weight := 0;

          If cbRight In Value Then
            Begin
              Border := Range.Borders[xlEdgeRight];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlEdgeRight].Weight := 0;

          If [cbLeft, cbRight] * Value <> cbsNone Then
            Begin
              Border := Range.Borders[xlInsideVertical];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlInsideVertical].Weight := 0;

          If cbTop In Value Then
            Begin
              Border := Range.Borders[xlEdgeTop];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlEdgeTop].Weight := 0;

          If cbBottom In Value Then
            Begin
              Border := Range.Borders[xlEdgeBottom];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlEdgeBottom].Weight := 0;

          If [cbTop, cbBottom] * Value <> cbsNone Then
            Begin
              Border := Range.Borders[xlInsideHorizontal];
              Border.Weight := xlThin;
              Border.LineStyle := xlContinuous;
            End
          Else
            Range.Borders[xlInsideHorizontal].Weight := 0;
        End
      Else If IsOpenOffice Then
        Begin
          Range := ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1);
          If Value <> cbsNone Then
            Begin
              BorderOn := ooCreateUnoStruct('com.sun.star.table.BorderLine');
              BorderOn.Color := 0;
              BorderOn.InnerLineWidth := 0;
              BorderOn.OuterLineWidth := 2;
              BorderOn.LineDistance := 0;
            End;
          If Value <> cbsAll Then
            Begin
              BorderOn := ooCreateUnoStruct('com.sun.star.table.BorderLine');
              BorderOn.Color := 0;
              BorderOn.InnerLineWidth := 0;
              BorderOn.OuterLineWidth := 0;
              BorderOn.LineDistance := 0;
            End;
          If cbLeft In Value Then
            Range.LeftBorder := BorderOn
          Else
            Range.LeftBorder := BOrderOff;
          If cbRight In Value Then
            Range.RightBorder := BorderOn
          Else
            Range.RightBorder := BOrderOff;
          If cbTop In Value Then
            Range.TopBorder := BorderOn
          Else
            Range.TopBorder := BOrderOff;
          If cbBottom In Value Then
            Range.BottomBorder := BorderOn
          Else
            Range.BottomBorder := BOrderOff;
        End;
    End;
End;

Procedure TSpreadSheet.SetRegionMerge(Row1, Col1, Row2, Col2: Integer;
  Value: Boolean);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          //xlRange(Row1, Col1, Row2, Col2).MergeCells := True;
          If Value Then
            xlRange(Row1, Col1, Row2, Col2).Merge
          Else
            xlRange(Row1, Col1, Row2, Col2).UnMerge;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).Merge(Value);
        End;
    End;
End;

Procedure TSpreadSheet.SetRegionProtected(Row1, Col1, Row2, Col2: Integer;
  Value: Boolean);
Var
  CellProtection: Variant;
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          xlRange(Row1, Col1, Row2, Col2).Locked := Value;
        End
      Else If IsOpenOffice Then
        Begin
          CellProtection := ooCreateUnoStruct('com.sun.star.util.CellProtection');
          CellProtection.IsLocked := Value;
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).CellProtection := CellProtection;
        End;
    End;
End;

Procedure TSpreadSheet.SetRegionWordWrap(Row1, Col1, Row2, Col2: Integer; Value: Boolean);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          xlRange(Row1, Col1, Row2, Col2).WrapText := Value;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).IsTextWrapped := Value;
        End;
    End;
End;

Procedure TSpreadSheet.SetRowHeight(Row, Height: Integer);
Begin
  If DocLoaded Then
    Begin
      If IsExcel Then
        Begin
          SheetSoftware.ActiveSheet.Rows[Row].RowHeight := Height * OOoToXlHeight;
        End
      Else If IsOpenOffice Then
        Begin
          ActiveSheet.getCellByPosition(0, Row - 1).getRows.getByIndex(0).Height := Height;
        End;
    End;

End;

// ***************************
// ** OpenOffice only stuff **
// ***************************

//Change 'C:\File.txt' into 'file:///c:/File.txt' (for OpenOffice OpenURL)

Function TSpreadSheet.FileName2URL(FileName: String): String;
Begin
  Result := '';
  If LowerCase(Copy(FileName, 1, 8)) <> 'file:///' Then
    Result := 'file:///';
  Result := Result + StringReplace(StringReplace(FileName, '\', '/', [rfReplaceAll, rfIgnoreCase]), ' ', '%20', [rfReplaceAll, rfIgnoreCase]);
End;

Function TSpreadSheet.ooCreateValue(ooName: String; ooData: Variant): Variant;
Var
  ooReflection: Variant;
Begin
  If IsOpenOffice Then
    Begin
      ooReflection := SheetSoftware.CreateInstance('com.sun.star.reflection.CoreReflection');
      ooReflection.forName('com.sun.star.beans.PropertyValue').createObject(Result);
      Result.Name := ooName;
      Result.Value := ooData;
    End
  Else
    Begin
      Raise Exception.Create('ooValue imposible to create, load OpenOffice first!');
    End;
End;

Function TSpreadSheet.ooCreateUnoStruct(StructName: String; IndexMax: Integer = -1): Variant;
Var
  I: Integer;
Begin
  Try
    If IndexMax < 0 Then
      Result := SheetSoftware.Bridge_GetStruct(StructName)
    Else
      Begin
        Result := VarArrayCreate([0, IndexMax], varVariant);
        For I := 0 To IndexMax Do
          Result[I] := SheetSoftware.Bridge_GetStruct(StructName);
      End;
  Except
    Result := Null;
  End;
  If VarIsEmpty(Result) Then
    Raise Exception.Create('Unknown structure name ' + StructName);
End;

Procedure TSpreadSheet.ooDispatch(ooCommand: String; ooParams: Variant);
Var
  ooDispatcher, ooFrame: Variant;
Begin
  If DocLoaded And IsOpenOffice Then
    Begin
      If (VarIsEmpty(ooParams) Or VarIsNull(ooParams)) Then
        ooParams := VarArrayCreate([0, -1], varVariant);
      ooFrame := Document.getCurrentController.getFrame;
      ooDispatcher := SheetSoftware.CreateInstance('com.sun.star.frame.DispatchHelper');
      ooDispatcher.executeDispatch(ooFrame, ooCommand, '', 0, ooParams);
    End
  Else
    Begin
      Raise Exception.Create('Dispatch imposible, load a OpenOffice doc first!');
    End;
End;

Function TSpreadSheet.ValidateSheetName(Name: String): String;
Begin
  Result := Name;
  If Trim(Result) = '' Then Result := 'Plan' + IntToStr(CountSheets);
  {If pos(':', Result) > 0 Then}Result := StringReplace(Result, ':', ' ', [rfReplaceAll]);
  {If pos('/', Result) > 0 Then}Result := StringReplace(Result, '/', ' ', [rfReplaceAll]);
  {If pos('?', Result) > 0 Then}Result := StringReplace(Result, '?', ' ', [rfReplaceAll]);
  {If pos('*', Result) > 0 Then}Result := StringReplace(Result, '*', ' ', [rfReplaceAll]);
  {If pos('[', Result) > 0 Then}Result := StringReplace(Result, '[', ' ', [rfReplaceAll]);
  {If pos(']', Result) > 0 Then}Result := StringReplace(Result, ']', ' ', [rfReplaceAll]);
  If Length(Result) > 31 Then Result := Copy(Result, 1, 31);
End;

// This function return coefficient for convartation Excel width to mm
// value depends of default font and are variable

Function TSpreadSheet.xlGetWidthMultiplier: Double;
// original was got from
// http://sources.codenet.ru/download/370/mdlSysMetrics.html
// this is translation from VisualBasic to Pascal with minimal modifications
Type
  enTextMetric = (
    entmHeight,
    entmInternalLeading,
    entmAveCharWidth,
    entmMaxCharWidth
    );

  Function CreateMyFont(sNameFace: String; nSize: Integer): Longint;
    //Create font handler
  Begin
    If Length(sNameFace) < 32 Then
      Result := CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDesktopWindow, LOGPIXELSY), 72), 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, PCHAR(sNameFace))
    Else
      Result := 0;
  End;

  Function MetricsFont(sNameFace: String; nSize: Integer; enType: enTextMetric): Longint;
    // return font metrics
    // if error return 0
    // sNameFace - installed font name, not greate then 31 char
    //             (examle: "Arial", "Times New Roman", etc)
    // nSize     - heightness in typographic point
    // enType    - metric's type, see enTextMetric

  Var
    hdc: Longint;
    hwnd: Longint;
    PrevMapMode: Longint;
    tm: TTextMetric;
    lFont: Longint;
    lOldFont: Longint;
  Begin
    Result := 0;
    Try
      hwnd := GetDesktopWindow;
      hdc := GetWindowDC(hwnd);
      If hdc <> 0 Then
        Begin
          PrevMapMode := SetMapMode(hdc, MM_LOMETRIC);
          lFont := CreateMyFont(sNameFace, nSize);
          If lFont <> 0 Then
            Begin
              lOldFont := SelectObject(hdc, lFont);
              GetTextMetrics(hdc, tm);
              Case enType Of
                entmHeight:
                  Result := tm.tmHeight;
                entmInternalLeading:
                  Result := tm.tmInternalLeading;
                entmAveCharWidth:
                  Result := tm.tmAveCharWidth;
                entmMaxCharWidth:
                  Result := tm.tmMaxCharWidth;
              End;
              SetMapMode(hdc, PrevMapMode);
              SelectObject(hdc, lOldFont);
              DeleteObject(lFont);
              ReleaseDC(hwnd, hdc);
            End;
        End
    Except
      Result := 0;
    End;
  End;

  Function WidthExcColToPixel(dblCountSymb: Double): Integer;
    // convert Excel's width units to pixels
    // if error return 0
  Var
    sNameFace: String;
    nSize: Integer;
    iAveCharWidth: Integer;
  Begin
    Result := 0;
    Try
      sNameFace := SheetSoftware.StandardFont;
      nSize := SheetSoftware.StandardFontSize;
      iAveCharWidth := MetricsFont(sNameFace, nSize, entmAveCharWidth);
      If iAveCharWidth > 0 Then
        Result := Round(iAveCharWidth * dblCountSymb) + 4;
    Except
      Result := 0;
    End;
  End;

Var
  W: Longint;
Begin
  W := MetricsFont(SheetSoftware.StandardFont, SheetSoftware.StandardFontSize, entmAveCharWidth);
  If W > 0 Then
    //Result := 72 / 25.4 / 100 / W
    Result := 0.1 / W //
  Else
    Result := 1 / 300; // if we can't get really value we will use approximate
End;

Function TSpreadSheet.xlRange(Row1, Col1, Row2, Col2: Integer): Variant;
Begin
  Result := ActiveSheet.Range[ActiveSheet.Cells[Row1, Col1], ActiveSheet.Cells[Row2, Col2]];
End;

Initialization

Finalization

End.

