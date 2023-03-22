unit frmMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComObj, variants;

{ TTemplate will be used to hold all fields to be replaced and their new values }
type
  TTemplate = record
    OTag: string;
    RTag: string;
  end;

type

  { TMainForm }

  TMainForm = class(TForm)
    btnCreateLibreDoc: TButton;
    btnReport: TButton;
    procedure btnCreateLibreDocClick(Sender: TObject);
    procedure btnReportClick(Sender: TObject);
  private

  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

{ TMainForm }

procedure TMainForm.btnCreateLibreDocClick(Sender: TObject);
// Do the same as within Delphi, Some changes needed to be made due to naming conflicts
const
  ServerName = 'com.sun.star.ServiceManager';
var
  Server: variant;
  Desktop, TextDocument, CoreReflection, HeaderRow, LoadParams, TextCursor, MyText,
  CellCursor, Table, Cell: variant;
  FontName: string;
  FileName: string;
  i, j: integer;
begin
  if Assigned(InitProc) then
    TProcedure(InitProc);

  // Try to create an instance of LibreOffice
  try
    Server := CreateOleObject(ServerName);
  except
    WriteLn('Unable to start LibreOffice!');
    Exit;
  end;

  // Desktop init
  Desktop := Server.CreateInstance('com.sun.star.frame.Desktop');
  CoreReflection := Server.CreateInstance(
    'com.sun.star.reflection.CoreReflection');

  // Initialize the Service Manager
  Server := CreateOleObject(ServerName);
  LoadParams := VarArrayCreate([0, -1], varVariant);
  TextDocument := Desktop.LoadComponentFromURL('private:factory/swriter',
    '_blank', 0, LoadParams);

  MyText := TextDocument.getText;
  // Assign a cursor
  TextCursor := MyText.createTextCursor;
    FontName := 'Arial';
  TextCursor.CharFontName := FontName;
  TextCursor.CharColor := $AAEECC;
  TextCursor.CharWeight := 150; // Make it bold
  MyText := TextDocument.Text;
  MyText.setString('Hello World of Lazarus');

  // Create an instance of the TextTable class
  Table := TextDocument.createInstance('com.sun.star.text.TextTable');
  Table.Initialize(8, 8);

  // Must insert the table first!
  MyText.insertTextContent(MyText.getEnd, Table, False);
  // Format the header, this can only be done after the previous line.
  HeaderRow := Table.GetRows.getByIndex(0);
  HeaderRow.BackColor := $AAEEAA;

  // Create a test table
  for i := 0 to 7 do
    for j := 0 to 7 do
    begin
      Cell := Table.getCellByPosition(i, j);
      // To change the color or the font of a cell, we need to create a textcursor object
      // for that given cell.
      CellCursor := Cell.getText.createTextCursor;
      CellCursor.CharFontName := 'Arial';

      Table.getCellByPosition(i, j).setString('Row ' + IntToStr(i + 1) +
        ', Column ' + IntToStr(j + 1));
    end;

  FileName := 'C:/temp/example.odt';
  TextDocument.storeToURL('file:///' + FileName,
    VarArrayCreate([0, 0], varVariant));

end;

procedure TMainForm.btnReportClick(Sender: TObject);
{ Replace a field in the template report }
var
   LOInstance, LOComponent, Textt, RText, LoadParams, TextBody, SearchDescriptor: Variant;
   newFileName, templateFileName: string;
begin

  templateFileName := 'file:///C:/temp/template.odt';
  newFileName := 'file:///C:/temp/Brief.odt';

  // Create instance of LibreOffice
  LoadParams := VarArrayCreate([0, -1], varVariant);
  LOInstance := CreateOleObject('com.sun.star.ServiceManager');
  LOComponent := LOInstance.createInstance('com.sun.star.frame.Desktop');

  // Load the template document
  TextBody := LOComponent.loadComponentFromURL(templateFileName, '_blank', 0, LoadParams);
  // Get the body of the document
  RText := TextBody.createReplaceDescriptor;
  RText.setSearchString('{NAME}');
  RText.setReplaceString('Danny Van Geyte');
  TextBody.ReplaceAll(RText);

  // Save the new file
  TextBody.storeToURL(newFileName, VarArrayCreate([0, 0], varVariant));

end;

end.
