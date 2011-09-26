// *******************************************************
// ** Delphi object for dual SpreadSheet managing using **
// ** Excel or OpenOffice in a transparent way.         **
// ** By: Sergio Hernandez (oficina(at)hcsoft.net)      **
// ** Version 1.00 30-06-2011 (DDMMYYYY)                **
// ** Use it freely, change it, etc. at will.           **
// ** Updates: Search for sergio-hcsoft in github:      **
// **      sergio-hcsoft/Delphi-SpreadSheets.git        **
// *******************************************************

{EXAMPLE OF USE
  //Create object: We have two flavours:
  //(A) from an existing file...
  HCalc:= THojaCalc.create(OpenDialog.FileName, false);
  //(B) from a blank document...
  HCalc:= THojaCalc.create(thcOpenOffice, true); //OpenOffice doc if possible, please
  HCalc.FileName:= 'C:\MyNewDoc'; //Needs a file name before you SaveDoc!
  //--end of creation.
  HCalc.ActivateSheetByIndex(2); //Activate second sheet
  if HCalc.IsActiveSheetProtected then
    ShowMessageHC('2nd sheet of name "'+HCalc.ActiveSheetName+'" IS protected');
  //Change a cell value.
  if HCalc.CellText[i,2] = '' then HCalc.CellText[i,2] := 'Hello world!';
  HCalc.AddNewSheet('New Sheet');
  HCalc.PrintDoc;
  HCalc.SaveDoc;
  HCalc.Free;
}

{TODO LIST:
  -PrintActiveSheet is not working for OpenOffice (is it even possible?)
  -Listener for OpenOffice so I can be notified if user visually close the doc.
}

{CHANGE LOG:
 V1.00:
   -Saving in Excel2007 will use Excel97 .xls file format instead of .xlsx
 V0.99:
   -Added a funtion by Alex Smith to set a cell text into italic.
 V0.98:
   -Added two procedures to easily send a number or a date to a cell position:
   SendDate(Row, Col, Date) and SendNumber(Row, Col, Float), if you look at
   the code you will notice that this is not so trivial as one could spect.
   -I have added (as comments) some useful code found on forums (copy-paste rows)
 V0.97:
   -Added CellFormula(col, row), similar to CellText, but allows to set a cell
   to a number wihout the efect of being considered by excel like a "text that
   looks like a number" (doesn't affect OpenOffice). Use it like this:
   CellFormula(1,1):= '=A2*23211.66';
   Note1: Excel will always spect numbers in this shape: no thousand separator
          and dot as decimal separator, regardless of your local configuration.
   Note2: Date is also bad interpreted in Excel, in this case you can use
          CellText but the date must be in american format: MM/DD/YYYY, if you
          use other format, it will try to interpret as an american date and
          only if it fails will use your local date format to "decode" it.
 V0.96:
   -Added PrintSheetsUntil(LastSheetName: string) -only works on excel- to print
   out all tabs from 1 until -excluded- the one with the given name in such a
   way that only one print job is created instead of one per tab (only way to do
   this in previous versions, so converting part of a excel to a single PDF
   using a printer like PDFCreator was not posible).
 V0.95:
   -ActivateSheetByIndex detect imposible index and allows to insert sheet 100 (it will create all necesary sheets)
   -SaveDocAs added a second optional parameter for OOo to use Excel97 format (rescued from V0.93 by Rômulo)
   -A little stronger ValidateSheetName() (filter away \ and " too).
 V0.94:
   -OpenOffice V2 compatible (small changes)
   -A lot of "try except" to avoid silly errors.
   -SaveDocAs(Name: string): boolean; (Added by Massimiliano Gozzi)
   -New function FileName2URL(Name) to convert from FileName to URL (OOo SaveDosAs)
   -New function ooCreateValue to hide all internals of OOo params creation
 V0.93:
   ***************************
   ** By Rômulo Silva Ramos **
   ***************************
   -FontSize(Row, Col, Size): change font size in that cell.
   -BackgroundColor(row, col: integer; color:TColor);
   -Add ValidateSheetName to validate sheet names when adding or renaming a sheet
   REVERTED FUNCTIONS (not neccesary in newer version V0.95 anymore)
   -Change AddNewSheet to add a new sheet in end at sheet list
   *REVERTED IN V0.95*
       It creates sheet following the active one, so to add at the end:
       ActivateSheetByIndex(CountSheets);
       AddNewSheet('Sheet '+IntToStr(CountSheets+1));
   -Change in SaveDoc to use SaveAs/StoreAsUrl
   *REVERTED V0.95*
       Use SaveDocAs(Name, true) for StoreAsUrl in Excel97 format.
 V0.92:
   -SetActiveSheetName didn't change the name to the right sheet on OpenOffice.
   -PrintPreview: New procedure to show up the print preview window.
   -Bold(Row, Col): Make bold the text in that cell.
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

{  PIECES OF CODE FOUND ON FORUMS WORTH COPYING HERE FOR FUTURE USE

  -Interesting "copy-paste one row to another" delphi code from PauLita posted
  on the OO forum (www.oooforum.org/forum/viewtopic.phtml?t=8878):

  OpenOffice version:
         Programa     := CreateOleObject( 'com.sun.star.ServiceManager' );
         ooParams     := VarArrayCreate([0,0],varVariant);
         ooParams[0]  := Programa.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
         ooView       := Document.getCurrentController;
         ooFrame      := ooView.getFrame;
         ooDispatcher := Programa.createInstance('com.sun.star.frame.DispatchHelper');
         // copy to clipboard
         oRange := Sheet.GetRows.GetByIndex(rl-1);
         ooView.Select( oRange );
         ooDispatcher.executeDispatch( ooFrame, '.uno:Copy',  '', 0, ooParams );
         // add one row to the table
         Sheet.GetRows.InsertByIndex(rl,1);
         // paste from clipboard
         oRange := Sheet.GetRows.GetByIndex(rl);
         ooView.Select( oRange );
         ooDispatcher.executeDispatch( ooFrame, '.uno:Paste',  '', 0, ooParams );
  Excel version:
         Sheet.Rows[r].Copy;
         Sheet.Rows[r+1].Insert(xlDown);
}

unit UHojaCalc;

interface

uses Variants, SysUtils, ComObj, Graphics, ActiveX;

//thcError: Tried to open but both failes
//thcNone:  Haven't tried still to open any
type TTipoHojaCalc = (thcError, thcNone, thcExcel, thcOpenOffice);

  TCellHorizontalAlign=(chaDefault, chaLeft, chaCenter, chaRight, chaBoth);

  TCellVerticalAlign=(cvaDefault, cvaTop, cvaCenter, cvaBottom);

  TCellBorder=(cbLeft, cbRight, cbTop, cbBottom);

  TCellBorders=Set Of TCellBorder;

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
  ooFileFormat: Array[TFileFormat] Of string = ('MS Excel 97');

  cbsNone: TCellBorders = [];
  cbsAll: TCellBorders = [cbLeft, cbRight, cbTop, cbBottom];

  // Excel mesure heightness in typographic point (1/72 inch), OOo hundredth of mm.
  OOoToXlHeight: Double = 72 / 25.4 / 100;
var
  //Excel mesure width in count of overage char width in defaukt font
  // this coefficient is variable and must be recalculated when default font is chanded
  OOoToXlWidth: Double;

type THojaCalc = class(TObject)
private
  fVisible:  boolean;
  //Program loaded stuff...
  procedure  LoadProg;
  procedure  CloseProg;
  function   GetProgLoaded: boolean;
  function   GetDocLoaded: boolean;
  function   GetIsExcel: boolean;
  function   GetIsOpenOffice: boolean;
  procedure  SetVisible(v: boolean);
  //Sheets stuff..
  function   GetCountSheets: integer;
  function   GetActiveSheetName: string;
  procedure  SetActiveSheetName(NewName: string);
  //Cells stuff...
  function   GetCellText(row,col: integer): string;
  procedure  SetCellText(row,col: integer; Txt: string);
  function   GetCellTextByName(Range: string): string;
  procedure  SetCellTextByName(Range: string; Txt: string);
  //OpenOffice only stuff...
  function   FileName2URL(FileName: string): string;
  procedure  ooDispatch(ooCommand: string; ooParams: variant);
  function   ooCreateValue(ooName: string; ooData: variant): variant;
  // Excel only stuff...
  function xlRange(Row1, Col1, Row2, Col2: integer): variant;
  function xlGetWidthMultiplier: Double;
  //Aux functions
  function   ValidateSheetName(Name:string): string;
public
  Tipo: TTipoHojaCalc;    //Witch program was used to manage the doc?
  FileName:    string;    //In windows FileName format C:\MyDoc.XXX
  Programa:    variant;   //Excel or OpenOfice instance created.
  DeskTop:     variant;   //OpenOffice desktop reference (not used now).
  Document:    variant;   //Document opened.
  ActiveSheet: variant;   //Active sheet.
  // TODO move to Private section
  function ooCreateUnoStruct(StructName: string; IndexMax: integer=-1): variant;
  //Object internals...
  constructor  Create(Name: string; MakeVisible: boolean); overload;
  constructor  Create(MyTipo: TTipoHojaCalc; MakeVisible: boolean); overload;
  destructor   Destroy; override;
  //Program loaded stuff...
  procedure    NewDoc;
  procedure    LoadDoc;
  procedure    CloseDoc;

  function     SaveDoc(FileFormat: TFileFormat=ffExcel97): boolean;
  function     PrintDoc: boolean;
  procedure    ShowPrintPreview;
  property     ProgLoaded: boolean     read GetProgLoaded;
  property     DocLoaded:  boolean     read GetDocLoaded;
  property     IsExcel: boolean        read GetIsExcel;
  property     IsOpenOffice: boolean   read GetIsOpenOffice;
  property     Visible: boolean        read fVisible           write SetVisible;
  //Sheets stuff...
  function     ActivateSheetByIndex(nIndex: integer): boolean;
  function     ActivateSheetByName(SheetName: string; CaseSensitive: boolean): boolean;
  function     IsActiveSheetProtected: boolean;
  function     ChangeActiveSheetProtection(IsProtected: boolean; Password: string): boolean;
  function     PrintActiveSheet: boolean;
  procedure    AddNewSheet(NewName: string);
  property     CountSheets:  integer   read GetCountSheets;
  property     ActiveSheetName: string read GetActiveSheetName write SetActiveSheetName;
  procedure    SetActiveSheetFont(Font: string);
  procedure    SetActiveSheetFontSize(Size: integer);
  // Region stuff...
  procedure    SetRegionMerge(Row1, Col1, Row2, Col2: integer; Value: boolean);
  procedure    SetRegionWordWrap(Row1, Col1, Row2, Col2: integer; Value: boolean);
  procedure    SetRegionBorder(Row1, Col1, Row2, Col2: integer; Value: TCellBorders);
  procedure    SetRegionProtected(Row1, Col1, Row2, Col2: integer; Value: boolean);
  //Cells stuff...
  procedure    SetCellFloat(row, col: integer; Value: Double);
  procedure    SetCellBold(row, col: integer; Value: boolean);
  procedure    SetCellBackgroundColor(row, col: integer; color: TColor);
  procedure    SetCellFontSize(row, col, Size: integer);
  procedure    SetCellFont(row, col: integer; Font: string);
  procedure    SetCellVerticalAlign(row, col: integer; Align: TCellVerticalAlign);
  procedure    SetCellHorizontalAlign(row, col: integer; Align: TCellHorizontalAlign);
  procedure    SetCellWordWrap(row, col: integer; Value: boolean);
  procedure    SetCellProtected(row, col: integer; Value: boolean);
  property     CellText[row, col: integer]: string read GetCellText write SetCellText;
  property     CellTextByName[Range: string]: string read GetCellTextByName write SetCellTextByName;
  //Other stuff
  procedure    SetColumnWidth(col, Width: integer); //Width in 1/100 of mm.
  procedure    SetRowHeight(row, Height: integer); //Height in 1/100 of mm.
end;

implementation

uses Windows;

// ************************
// ** Create and destroy **
// ************************

//Create with an empty doc of requested type (use thcExcel or thcOpenOffice)
//Remember to define FileName before calling to SaveDoc
constructor THojaCalc.Create(MyTipo: TTipoHojaCalc; MakeVisible: boolean);
var
  i: integer;
  IsFirstTry: boolean;
begin
  //Close all opened things first...
  if DocLoaded then CloseDoc;
  if ProgLoaded then CloseProg;
  //I will try to open twice, so if Excel fails, OpenOffice is used instead
  IsFirstTry:= true;
  for i:= 1 to 2 do begin
    //Try to open Excel...
    if (MyTipo = thcExcel) or (MyTipo = thcNone) then begin
      
        Programa:= CreateOleObject('Excel.Application');
      
      
      if ProgLoaded then begin
        Tipo:= thcExcel;
        break;
      end else begin
        if IsFirstTry then begin
          //Try OpenOffice as my second choice
          MyTipo:= thcOpenOffice;
          IsFirstTry:= false;
        end else begin
          //Both failed!
          break;
        end;
      end;
    end;
    //Try to open OpenOffice...
    if (MyTipo = thcOpenOffice) or (MyTipo = thcNone)then begin
      try
        Programa:= CreateOleObject('com.sun.star.ServiceManager');
      
      
      if ProgLoaded then begin
        Tipo:= thcOpenOffice;
        break;
      end else begin
        if IsFirstTry then begin
          //Try Excel as my second choice
          MyTipo:= thcExcel;
          IsFirstTry:= false;
        end else begin
          //Both failed!
          break;
        end;
      end;
    end;
  end;
  //Was it able to open any of them?
  if Tipo = thcNone then begin
    Tipo:= thcError;
    raise Exception.Create('THojaCalc.create failed, may be no Office is installed?');
  end;
  //Add a blank document...
  fVisible:= MakeVisible;
  NewDoc;
end;

constructor THojaCalc.Create(Name: string; MakeVisible: boolean);
begin
  //Store values...
  FileName:= Name;
  fVisible:=  MakeVisible;
  //Open program and document...
  LoadProg;
  LoadDoc;
end;

destructor THojaCalc.Destroy;
begin

    CloseDoc;
    CloseProg;
  inherited;
end;

// *************************
// ** Loading the program **
// ** Excel or OpenOffice **
// *************************

procedure THojaCalc.LoadProg;
begin
  if ProgLoaded then CloseProg;
  if (UpperCase(ExtractFileExt(FileName))='.XLS') then begin
    //Excel is the primary choice...
    
      Programa:= CreateOleObject('Excel.Application');
    
    if ProgLoaded then Tipo:= thcExcel;
  end;
  //Not lucky with Excel? Another filetype? Let's go with OpenOffice...
  if Tipo = thcNone then begin
    //Try with OpenOffice...
    
      Programa:= CreateOleObject('com.sun.star.ServiceManager');
    
    if ProgLoaded then Tipo:= thcOpenOffice;
  end;
  //Still no program loaded?
  if not ProgLoaded then begin
    Tipo:= thcError;
    raise Exception.Create('THojaCalc.create failed, may be no Office is installed?');
  end;
end;

procedure THojaCalc.CloseProg;
begin
  if DocLoaded then CloseDoc;
  if ProgLoaded then begin
    try
      if IsExcel then      Programa.Quit;
      //Next line made OO V2 not to work anymore as the next call to
      //CreateOleObject('com.sun.star.ServiceManager') failed.
      //if IsOpenOffice then Programa.Dispose;
      Programa:= Unassigned;
    finally end;
  end;
  Tipo:= thcNone;
end;

//Is there any prog loaded? Witch one?
function THojaCalc.GetProgLoaded: boolean;
begin
  result:= not (VarIsEmpty(Programa) or VarIsNull(Programa));
end;
function  THojaCalc.GetIsExcel: boolean;
begin
  result:= (Tipo=thcExcel);
end;
function  THojaCalc.GetIsOpenOffice: boolean;
begin
  result:= (Tipo=thcOpenOffice);
end;

// ************************
// ** Loading a document **
// ************************

procedure THojaCalc.NewDoc;
var ooParams: variant;
begin
  //Is the program running? (Excel or OpenOffice)
  if not ProgLoaded then raise Exception.Create('No program loaded for the new document.');
  //Is there a doc already loaded?
  if DocLoaded then CloseDoc;
  DeskTop:= Unassigned;
  //OK, now try to create the doc...
  if IsExcel then begin
    Programa.WorkBooks.Add;
    Programa.Visible:= Visible;
    Programa.DisplayAlerts:= false;
    Document:= Programa.ActiveWorkBook;
    ActiveSheet:= Document.ActiveSheet;
    OOoToXlWidth:= xlGetWidthMultiplier;
  end
  else if IsOpenOffice then begin
    Desktop:= Programa.CreateInstance('com.sun.star.frame.Desktop');
    //Optional parameters (visible)...
    ooParams:=    VarArrayCreate([0, 0], varVariant);
    ooParams[0]:= ooCreateValue('Hidden', not Visible);
    //Create the document...
    Document:= Desktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, ooParams);
    ActivateSheetByIndex(1);
  end;
end;

procedure THojaCalc.LoadDoc;
var ooParams: variant;
begin
  if FileName='' then exit;
  //Is the program running? (Excel or OpenOffice)
  if not ProgLoaded then LoadProg;
  //Is there a doc already loaded?
  if DocLoaded then CloseDoc;
  DeskTop:= Unassigned;
  //OK, now try to open the doc...
  if IsExcel then begin
    Programa.WorkBooks.Open(FileName, 3);
    Programa.Visible:= Visible;
    Programa.DisplayAlerts:= false;
    Document:= Programa.ActiveWorkBook;
    ActiveSheet:= Document.ActiveSheet;
    OOoToXlWidth:= xlGetWidthMultiplier;
  end
  else if IsOpenOffice then begin
    Desktop:= Programa.CreateInstance('com.sun.star.frame.Desktop');
    //Optional parameters (visible)...
    ooParams:=    VarArrayCreate([0, 0], varVariant);
    //Next line stop working OK on OOo V2: Created blind, always blind!
    //so now it is create as visible, then set to non visible if requested
    ooParams[0]:= ooCreateValue('Hidden', not Visible);
    //Open the document...
    Document:= Desktop.LoadComponentFromURL(FileName2URL(FileName), '_blank', 0, ooParams);
    ActivateSheetByIndex(1);
  end;
  if Tipo=thcNone then
    raise Exception.Create('not able to read the file "' + FileName + '". Probably not exist a software installed to open it.');
end;

function THojaCalc.SaveDoc(FileFormat: TFileFormat=ffExcel97): boolean;
var
  ooParams: variant;
  Dir: string;
begin
  result:= false;
  Dir:= ExtractFilePath(FileName);
  if not DirectoryExists(Dir) then ForceDirectories(Dir);
  if DocLoaded then begin
    if IsExcel then begin
      Document.SaveAs(FileName, XlFileFormat[FileFormat]);
      result:= true;
    end
    else if IsOpenOffice then begin
      ooParams:= VarArrayCreate([0, 0], varVariant);
      ooParams[0]:= ooCreateValue('FilterName', ooFileFormat[FileFormat]);
      Document.StoreAsUrl(FileName2URL(FileName), ooParams);
      result:= true;
    end;
  end;
end;


//Print the Doc...
function THojaCalc.PrintDoc: boolean;
var ooParams: variant;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      Document.PrintOut;
      result:= true;
    end
    else if IsOpenOffice then begin
      //NOTE: OpenOffice will print all sheets with Printable areas, but if no
      //printable areas are defined in the doc, it will print all entire sheets.
      //Optional parameters (wait until fully sent to printer)...
      ooParams:=   VarArrayCreate([0, 0], varVariant);
      ooParams[0]:= ooCreateValue('Wait', true);
      Document.Print(ooParams);
      result:= true;
    end;
  end;
end;

procedure THojaCalc.ShowPrintPreview;
begin
  if DocLoaded then begin
    //Force visibility of the doc...
    Visible:= true;
    if IsExcel then begin // TODO replase first 3 parameters by default value instead of null
      Document.PrintOut(Null, Null, Null, true);
    end
    else if IsOpenOffice then begin
      ooDispatch('.uno:PrintPreview', Unassigned);
    end;
  end;
end;

procedure THojaCalc.SetVisible(v: boolean);
begin
  if DocLoaded and (v<>fVisible) then begin
    if IsExcel then begin
      Programa.Visible:= v;
    end
    else if IsOpenOffice then begin
      Document.getCurrentController.getFrame.getContainerWindow.SetVisible(v);
    end;
    fVisible:= v;
  end;
end;

procedure THojaCalc.CloseDoc;
begin
  if DocLoaded then begin
    //Close it...
    try
      if IsOpenOffice then begin
        Document.Dispose;
      end
      else if IsExcel then begin
        Document.Close;
      end;
    finally end;
    //Clean up both "pointer"...
    Document:= Null;
    ActiveSheet:= Null;
  end;
end;

function THojaCalc.GetDocLoaded: boolean;
begin
  result:= not (VarIsEmpty(Document) or VarIsNull(Document));
end;

// *********************
// ** Managing sheets **
// *********************

function THojaCalc.GetCountSheets: integer;
begin
  result:= 0;
  if DocLoaded then begin
    if IsExcel then begin
      result:= Document.Sheets.Count;
    end
    else if IsOpenOffice then begin
      result:= Document.getSheets.GetCount;
    end;
  end;
end;

//Index is 1 based in Excel, but OpenOffice uses it 0-based
//Here we asume 1-based so OO needs to activate (nIndex-1)
function THojaCalc.ActivateSheetByIndex(nIndex: integer): boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      Document.Sheets[nIndex].activate;
      ActiveSheet:= Document.ActiveSheet;
      result:= true;
    end
    else if IsOpenOffice then begin
      ActiveSheet:= Document.getSheets.getByIndex(nIndex - 1);
      result:= true;
    end;
    sleep(100); //Asyncronus, so better give it time to make the change
  end;
end;

//Find a sheet by its name...
function THojaCalc.ActivateSheetByName(SheetName: string; CaseSensitive: boolean): boolean;
var
  OldActiveSheet: variant;
  i: integer;
begin
  result:= false;
  if DocLoaded then begin
    if CaseSensitive then begin
      //Find the EXACT name...
      if IsExcel then begin
        Document.Sheets[SheetName].Select;
        ActiveSheet:= Document.ActiveSheet;
        result:= true;
      end
      else if IsOpenOffice then begin
        ActiveSheet:= Document.getSheets.getByName(SheetName);
        result:= true;
      end;
    end else begin
      //Find the Sheet regardless of the case...
      OldActiveSheet:= ActiveSheet;
      for i:= 1 to GetCountSheets do begin
        ActivateSheetByIndex(i);
        if UpperCase(ActiveSheetName)=UpperCase(SheetName) then begin
          result:= true;
          Exit;
        end;
      end;
      //If not found, let the old active sheet active...
      ActiveSheet:= OldActiveSheet;
    end;
  end;
end;

//Name of the active sheet?
function THojaCalc.GetActiveSheetName: string;
begin
  if DocLoaded then begin
    if IsExcel then begin
      result:= ActiveSheet.Name;
    end
    else if IsOpenOffice then begin
      result:= ActiveSheet.GetName;
    end;
  end;
end;

procedure THojaCalc.SetActiveSheetFont(Font: string);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells.Font.Name:= Font;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.CharFontName:= Font;
        end;
    end;
end;

procedure THojaCalc.SetActiveSheetFontSize(Size: integer);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells.Font.Size:= Size;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.CharHeight:= Size;
        end;
    end;
end;

procedure THojaCalc.SetActiveSheetName(NewName: string);
begin
  NewName:= ValidateSheetName(NewName);
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Name:= NewName;
    end
    else if IsOpenOffice then begin
      ActiveSheet.setName(NewName);
      //This code always changes the name of "visible" sheet, not active one!
      //ooParams:= VarArrayCreate([0, 0], varVariant);
      //ooParams[0]:= ooCreateValue('Name', NewName);
      //ooDispatch('.uno:RenameTable', ooParams);
    end;
  end;
end;

function THojaCalc.ChangeActiveSheetProtection(IsProtected: boolean;
  Password: string): boolean;
begin
  if DocLoaded then begin
    if IsExcel then begin
      if IsProtected then
        ActiveSheet._Protect(Password)
      else
        ActiveSheet.Unprotect(Password);
    end
    else if IsOpenOffice then begin
      if IsProtected then
        ActiveSheet.Protect(Password)
      else
        ActiveSheet.Unprotect(Password);
    end;
  end;
  result:= IsActiveSheetProtected;
end;

//Check for sheet protection (password)...
function THojaCalc.IsActiveSheetProtected: boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      result:= ActiveSheet.ProtectContents;
    end
    else if IsOpenOffice then begin
      result:= ActiveSheet.IsProtected;
    end;
  end;
end;

//WARNING: This function is NOT dual, only works for Excel docs!
//Send active sheet to default printer (as seen in preview window)...
function THojaCalc.PrintActiveSheet: boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      ActiveSheet.PrintOut;
      result:= true;
    end
    else if IsOpenOffice then begin
      raise Exception.Create('Function "PrintActiveSheet" still not working in OpenOffice!');
      //ActiveSheet.Print;
      result:= false;
    end;
  end;
end;


//Add a new sheet, name it, and make it the active sheet...
procedure THojaCalc.AddNewSheet(NewName: string);
var
  ooSheets: variant;
begin
  NewName := ValidateSheetName(NewName);
  if DocLoaded then begin
    if IsExcel then begin
      Document.WorkSheets.Add;
      Document.ActiveSheet.Move(After:= Document.Sheets[CountSheets]);
      Document.ActiveSheet.Name:= NewName;
      //Active sheet has move to this new one, so I need to update the var
      ActiveSheet:= Document.ActiveSheet;
    end
    else if IsOpenOffice then begin
      ooSheets:= Document.getSheets;
      ooSheets.insertNewByName(NewName, CountSheets);
      //Redefine active sheet to this new one
      ActiveSheet:= ooSheets.getByName(NewName);
    end;
  end;
end;

function THojaCalc.ValidateSheetName(Name: string): string;
begin
  result:= Name;
  if Trim(result)='' then result:= 'Plan' + IntToStr(CountSheets);
  result:= StringReplace(result, ':', ' ', [rfReplaceAll]);
  result:= StringReplace(result, '/', ' ', [rfReplaceAll]);
  result:= StringReplace(result, '?', ' ', [rfReplaceAll]);
  result:= StringReplace(result, '*', ' ', [rfReplaceAll]);
  result:= StringReplace(result, '[', ' ', [rfReplaceAll]);
  result:= StringReplace(result, ']', ' ', [rfReplaceAll]);
  if Length(result) > 31 then result:= Copy(result, 1, 31);
end;

// ************************
// ** Manage  the  cells **
// ** in the ActiveSheet **
// ************************

//Read/Write cell text (formula en Excel) by index
//OpenOffice start at cell (0,0) while Excel at (1,1)
//Also, Excel uses (row, col) and OpenOffice uses (col, row)
function THojaCalc.GetCellText(row, col: integer): string;
begin
  if DocLoaded then begin
    if IsExcel then begin
      result:= ActiveSheet.Cells[row, col].Formula; //.Text;
    end else
    if IsOpenOffice then begin
      result:= ActiveSheet.getCellByPosition(col-1, row-1).getFormula;
    end;
  end;
end;
procedure  THojaCalc.SetCellText(row, col: integer; Txt: string);
begin
  if DocLoaded then begin
    if IsExcel then begin
      ActiveSheet.Cells[row, col].Formula:= Txt;
    end
    else if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).setFormula(Txt);
    end;
  end;
end;

//read/write cell text (formula in excel) by name instead of position
//for instance, you can set the value for cell 'NewSheet!A12' or similar
//NOTE: if range contains several cells, first one will be used.

function THojaCalc.GetCellTextByName(Range: string): string;
var
  OldActiveSheet: variant;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          result:= Programa.Range[Range].Text; //Set 'Formula' but Get 'Text';
        end
      else if IsOpenOffice then
        begin
          OldActiveSheet:= ActiveSheet;
          //if range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
          if pos('!', Range) > 0 then
            begin
              //Activate the proper sheet...
              if not ActivateSheetByName(Copy(Range, 1, pos('!', Range) - 1), false) then
                raise Exception.Create('Sheet "' + Copy(Range, 1, pos('!', Range) - 1) + '" not present in the document.');
              Range:= Copy(Range, pos('!', Range) + 1, 999);
            end;
          result:= ActiveSheet.getCellRangeByName(Range).getCellByPosition(0, 0).getFormula;
          ActiveSheet:= OldActiveSheet;
        end;
    end;
end;

procedure THojaCalc.SetCellTextByName(Range: string; Value: string);
var
  OldActiveSheet: variant;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.Range[Range].Formula:= Value;
        end
      else if IsOpenOffice then
        begin
          OldActiveSheet:= ActiveSheet;
          //if range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
          if pos('!', Range) > 0 then
            begin
              //Activate the proper sheet...
              if not ActivateSheetByName(Copy(Range, 1, pos('!', Range) - 1), false) then
                raise Exception.Create('Sheet "' + Copy(Range, 1, pos('!', Range) - 1) + '" not present in the document.');
              Range:= Copy(Range, pos('!', Range) + 1, 999);
            end;
          ActiveSheet.getCellRangeByName(Range).getCellByPosition(0, 0).setFormula(Value);
          ActiveSheet:= OldActiveSheet;
        end;
    end;
end;

procedure THojaCalc.SetCellFloat(row, col: integer; Value: Double);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          ActiveSheet.Cells[row, col].Formula:= Value;
        end
      else if IsOpenOffice then
        begin
          if IsOpenOffice then ActiveSheet.getCellByPosition(col - 1, row - 1).setValue(Value);
        end;
    end;
end;

procedure THojaCalc.SetCellVerticalAlign(row, col: integer;
  Align: TCellVerticalAlign);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].VerticalAlignment:= XlVAlign[Align];
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).VertJustify:= ooCellVertJustify[Align];
        end;
    end;
end;

procedure THojaCalc.SetCellWordWrap(row, col: integer; Value: boolean);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].WrapText:= Value;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).IsTextWrapped:= Value;
        end;
    end;
end;

procedure THojaCalc.SetCellBackgroundColor(row, col: integer; Color: TColor);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].Interior.Color:= Color;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).CellBackColor:= Color;
        end;
    end;
end;

procedure THojaCalc.SetCellFont(row, col: integer; Font: string);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].Font.Name:= Font;
        end
      else if
        IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).CharFontName:= Font;
        end;
    end;
end;

procedure THojaCalc.SetCellFontSize(row, col, Size: integer);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].Font.Size:= Size;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).CharHeight:= Size;
        end;
    end;
end;

procedure THojaCalc.SetCellHorizontalAlign(row, col: integer;
  Align: TCellHorizontalAlign);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].HorizontalAlignment:= XlHAlign[Align];
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(col - 1, row - 1).HoriJustify:= ooCellHoriJustify[Align];
        end;
    end;
end;

procedure THojaCalc.SetCellProtected(row, col: integer; Value: boolean);
var
  CellProtection: variant;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          if IsExcel then ActiveSheet.Cells[row, col].Locked:= Value;
        end
      else if IsOpenOffice then
        begin
          CellProtection:= ooCreateUnoStruct('com.sun.star.util.CellProtection');
          CellProtection.IsLocked:= Value;
          ActiveSheet.getCellByPosition(col - 1, row - 1).CellProtection:= CellProtection;
        end;
    end;
end;

procedure THojaCalc.SetCellBold(row, col: integer; Value: boolean);
Const
  //OpenOffice.org_3.0_SDK\sdk\idl\com\sun\star\awt\FontWeight.idl
  ooBold: integer=150; //com.sun.star.awt.FontWeight.BOLD
  ooNorman: integer=100; //com.sun.star.awt.FontWeight.NORMAL
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Cells[row, col].Font.Bold:= Value;
        end
      else if IsOpenOffice then
        begin
          if Value then
            ActiveSheet.getCellByPosition(col - 1, row - 1).CharWeight:= ooBold
          else
            ActiveSheet.getCellByPosition(col - 1, row - 1).CharWeight:= ooNorman;
        end;
    end;
end;

procedure THojaCalc.SetColumnWidth(col, Width: integer); //Width in 1/100 of mm.
var
  Column: variant;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          //Excel use the width of '0' as the unit, we do an aproximation: Width '0'=2 mm.
          { TODO : find how to get exactly value of this const}

          Programa.ActiveSheet.Columns[col].ColumnWidth:= Width * OOoToXlWidth;
        end
      else if IsOpenOffice then
        begin
          Column:= ActiveSheet.getCellByPosition(col - 1, 0).getColumns.getByIndex(0);
          Column.Width:= Width;
          Column.IsVisible:= (Width > 0);
        end;
    end;
end;

procedure THojaCalc.SetRegionBorder(Row1, Col1, Row2, Col2: integer;
  Value: TCellBorders);
var
  Range, BorderOn, BOrderOff, Border: variant;
Const
  xlThin=$00000002;
  xlContinuous=$00000001;
  xlInsideHorizontal=$0000000C;
  xlInsideVertical=$0000000B;
  xlDiagonalDown=$00000005;
  xlDiagonalUp=$00000006;
  xlEdgeBottom=$00000009;
  xlEdgeLeft=$00000007;
  xlEdgeRight=$0000000A;
  xlEdgeTop=$00000008;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Range:= xlRange(Row1, Col1, Row2, Col2);

          if cbLeft In Value then
            begin
              Border:= Range.Borders[xlEdgeLeft];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlEdgeLeft].Weight:= 0;

          if cbRight In Value then
            begin
              Border:= Range.Borders[xlEdgeRight];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlEdgeRight].Weight:= 0;

          if [cbLeft, cbRight] * Value <> cbsNone then
            begin
              Border:= Range.Borders[xlInsideVertical];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlInsideVertical].Weight:= 0;

          if cbTop In Value then
            begin
              Border:= Range.Borders[xlEdgeTop];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlEdgeTop].Weight:= 0;

          if cbBottom In Value then
            begin
              Border:= Range.Borders[xlEdgeBottom];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlEdgeBottom].Weight:= 0;

          if [cbTop, cbBottom] * Value <> cbsNone then
            begin
              Border:= Range.Borders[xlInsideHorizontal];
              Border.Weight:= xlThin;
              Border.LineStyle:= xlContinuous;
            end
          else
            Range.Borders[xlInsideHorizontal].Weight:= 0;
        end
      else if IsOpenOffice then
        begin
          Range:= ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1);
          if Value <> cbsNone then
            begin
              BorderOn:= ooCreateUnoStruct('com.sun.star.table.BorderLine');
              BorderOn.Color:= 0;
              BorderOn.InnerLineWidth:= 0;
              BorderOn.OuterLineWidth:= 2;
              BorderOn.LineDistance:= 0;
            end;
          if Value <> cbsAll then
            begin
              BorderOn:= ooCreateUnoStruct('com.sun.star.table.BorderLine');
              BorderOn.Color:= 0;
              BorderOn.InnerLineWidth:= 0;
              BorderOn.OuterLineWidth:= 0;
              BorderOn.LineDistance:= 0;
            end;
          if cbLeft In Value then
            Range.LeftBorder:= BorderOn
          else
            Range.LeftBorder:= BOrderOff;
          if cbRight In Value then
            Range.RightBorder:= BorderOn
          else
            Range.RightBorder:= BOrderOff;
          if cbTop In Value then
            Range.TopBorder:= BorderOn
          else
            Range.TopBorder:= BOrderOff;
          if cbBottom In Value then
            Range.BottomBorder:= BorderOn
          else
            Range.BottomBorder:= BOrderOff;
        end;
    end;
end;

procedure THojaCalc.SetRegionMerge(Row1, Col1, Row2, Col2: integer;
  Value: boolean);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          //xlRange(Row1, Col1, Row2, Col2).MergeCells:= true;
          if Value then
            xlRange(Row1, Col1, Row2, Col2).Merge
          else
            xlRange(Row1, Col1, Row2, Col2).UnMerge;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).Merge(Value);
        end;
    end;
end;

procedure THojaCalc.SetRegionProtected(Row1, Col1, Row2, Col2: integer;
  Value: boolean);
var
  CellProtection: variant;
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          xlRange(Row1, Col1, Row2, Col2).Locked:= Value;
        end
      else if IsOpenOffice then
        begin
          CellProtection:= ooCreateUnoStruct('com.sun.star.util.CellProtection');
          CellProtection.IsLocked:= Value;
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).CellProtection:= CellProtection;
        end;
    end;
end;

procedure THojaCalc.SetRegionWordWrap(Row1, Col1, Row2, Col2: integer; Value: boolean);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          xlRange(Row1, Col1, Row2, Col2).WrapText:= Value;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellRangeByPosition(Col1 - 1, Row1 - 1, Col2 - 1, Row2 - 1).IsTextWrapped:= Value;
        end;
    end;
end;

procedure THojaCalc.SetRowHeight(row, Height: integer);
begin
  if DocLoaded then
    begin
      if IsExcel then
        begin
          Programa.ActiveSheet.Rows[row].RowHeight:= Height * OOoToXlHeight;
        end
      else if IsOpenOffice then
        begin
          ActiveSheet.getCellByPosition(0, row - 1).getRows.getByIndex(0).Height:= Height;
        end;
    end;

end;

// ***************************
// ** OpenOffice only stuff **
// ***************************

//Change 'C:\File.txt' into 'file:///c:/File.txt' (for OpenOffice OpenURL)

function THojaCalc.FileName2URL(FileName: string): string;
begin
  result:= '';
  if LowerCase(Copy(FileName, 1, 8)) <> 'file:///' then
    result:= 'file:///';
  result:= result + StringReplace(StringReplace(FileName, '\', '/', [rfReplaceAll, rfIgnoreCase]), ' ', '%20', [rfReplaceAll, rfIgnoreCase]);
end;

function THojaCalc.ooCreateValue(ooName: string; ooData: variant): variant;
var
  ooReflection: variant;
begin
  if IsOpenOffice then
    begin
      ooReflection:= Programa.CreateInstance('com.sun.star.reflection.CoreReflection');
      ooReflection.forName('com.sun.star.beans.PropertyValue').createObject(result);
      result.Name:= ooName;
      result.Value:= ooData;
    end
  else
    begin
      raise Exception.Create('ooValue imposible to create, load OpenOffice first!');
    end;
end;

function THojaCalc.ooCreateUnoStruct(StructName: string; IndexMax: integer=-1): variant;
var
  I: integer;
begin
  try
    if IndexMax < 0 then
      result:= Programa.Bridge_GetStruct(StructName)
    else
      begin
        result:= VarArrayCreate([0, IndexMax], varVariant);
        for I:= 0 to IndexMax do
          result[I]:= Programa.Bridge_GetStruct(StructName);
      end;
  except
    result:= Null;
  end;
  if VarIsEmpty(result) then
    raise Exception.Create('Unknown structure name ' + StructName);
end;

procedure THojaCalc.ooDispatch(ooCommand: string; ooParams: variant);
var
  ooDispatcher, ooFrame: variant;
begin
  if DocLoaded And IsOpenOffice then
    begin
      if (VarIsEmpty(ooParams) or VarIsNull(ooParams)) then
        ooParams:= VarArrayCreate([0, -1], varVariant);
      ooFrame:= Document.getCurrentController.getFrame;
      ooDispatcher:= Programa.CreateInstance('com.sun.star.frame.DispatchHelper');
      ooDispatcher.executeDispatch(ooFrame, ooCommand, '', 0, ooParams);
    end
  else
    begin
      raise Exception.Create('Dispatch imposible, load a OpenOffice doc first!');
    end;
end;

// This function return coefficient for convartation Excel width to mm
// value depends of default font and are variable

function THojaCalc.xlGetWidthMultiplier: Double;
// original was got from
// http://sources.codenet.ru/download/370/mdlSysMetrics.html
// this is translation from VisualBasic to Pascal with minimal modifications
Type
  enTextMetric=(
    entmHeight,
    entmInternalLeading,
    entmAveCharWidth,
    entmMaxCharWidth
    );

  function CreateMyFont(sNameFace: string; nSize: integer): Longint;
    //Create font handler
  begin
    if Length(sNameFace) < 32 then
      result:= CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDesktopWindow, LOGPIXELSY), 72), 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, PCHAR(sNameFace))
    else
      result:= 0;
  end;

  function MetricsFont(sNameFace: string; nSize: integer; enType: enTextMetric): Longint;
    // return font metrics
    // if error return 0
    // sNameFace - installed font name, not greate then 31 char
    //             (examle: "Arial", "Times New Roman", etc)
    // nSize     - heightness in typographic point
    // enType    - metric's type, see enTextMetric

  var
    hdc: Longint;
    hwnd: Longint;
    PrevMapMode: Longint;
    tm: TTextMetric;
    lFont: Longint;
    lOldFont: Longint;
  begin
    result:= 0;
    try
      hwnd:= GetDesktopWindow;
      hdc:= GetWindowDC(hwnd);
      if hdc <> 0 then
        begin
          PrevMapMode:= SetMapMode(hdc, MM_LOMETRIC);
          lFont:= CreateMyFont(sNameFace, nSize);
          if lFont <> 0 then
            begin
              lOldFont:= SelectObject(hdc, lFont);
              GetTextMetrics(hdc, tm);
              Case enType Of
                entmHeight:
                  result:= tm.tmHeight;
                entmInternalLeading:
                  result:= tm.tmInternalLeading;
                entmAveCharWidth:
                  result:= tm.tmAveCharWidth;
                entmMaxCharWidth:
                  result:= tm.tmMaxCharWidth;
              end;
              SetMapMode(hdc, PrevMapMode);
              SelectObject(hdc, lOldFont);
              DeleteObject(lFont);
              ReleaseDC(hwnd, hdc);
            end;
        end
    except
      result:= 0;
    end;
  end;

  function WidthExcColToPixel(dblCountSymb: Double): integer;
    // convert Excel's width units to pixels
    // if error return 0
  var
    sNameFace: string;
    nSize: integer;
    iAveCharWidth: integer;
  begin
    result:= 0;
    try
      sNameFace:= Programa.StandardFont;
      nSize:= Programa.StandardFontSize;
      iAveCharWidth:= MetricsFont(sNameFace, nSize, entmAveCharWidth);
      if iAveCharWidth > 0 then
        result:= Round(iAveCharWidth * dblCountSymb) + 4;
    except
      result:= 0;
    end;
  end;

var
  W: Longint;
begin
  W:= MetricsFont(Programa.StandardFont, Programa.StandardFontSize, entmAveCharWidth);
  if W > 0 then
    //result:= 72 / 25.4 / 100 / W
    result:= 0.1 / W //
  else
    result:= 1 / 300; // if we can't get really value we will use approximate
end;

function THojaCalc.xlRange(Row1, Col1, Row2, Col2: integer): variant;
begin
  result:= ActiveSheet.Range[ActiveSheet.Cells[Row1, Col1], ActiveSheet.Cells[Row2, Col2]];
end;

Initialization

Finalization

end.

