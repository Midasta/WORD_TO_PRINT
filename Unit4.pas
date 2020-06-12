unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.StdCtrls, Vcl.Mask, Vcl.ExtCtrls, Data.Win.ADODB,comobj,ShellApi;

type
  TForm4 = class(TForm)
    Panel1: TPanel;
    ComboBox1: TComboBox;
    Label1: TLabel;
    LabeledEdit1: TLabeledEdit;
    Label2: TLabel;
    MaskEdit1: TMaskEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    Panel2: TPanel;
    DBGrid1: TDBGrid;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    ADOQuery1NAME_RU: TWideStringField;
    ADOQuery1POL: TWideStringField;
    ADOQuery1DOB: TDateTimeField;
    ADOQuery1ADRESS: TWideStringField;
    ADOQuery1SPEC: TWideStringField;
    Panel3: TPanel;
    Button3: TButton;
    Button4: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;
  Excel,WordApp,FData: Variant;
  i,j:integer;
  const
   wdReplaceAll=2;

implementation

{$R *.dfm}



procedure TForm4.Button1Click(Sender: TObject);
begin
ADOQuery1.Insert;
ADOQuery1NAME_RU.Value:=LabeledEdit1.Text;
ADOQuery1POL.Value:=ComboBox1.Text;
ADOQuery1DOB.Value:=strtodate(MaskEdit1.Text);
ADOQuery1ADRESS.Value:=LabeledEdit2.Text;
ADOQuery1SPEC.Value:=LabeledEdit3.Text;
ADOQuery1.Next;
ADOQuery1.Active:=false;
ADOQuery1.Active:=true;
Button2Click(Self);
end;

procedure TForm4.Button2Click(Sender: TObject);
begin
LabeledEdit1.Clear;
ComboBox1.Clear;
MaskEdit1.Clear;
LabeledEdit2.Clear;
LabeledEdit3.Clear;
end;

procedure TForm4.Button3Click(Sender: TObject);
begin

 for i := 0 to ADOQuery1.RecordCount-1 do
 begin
    try
     Wordapp:=CreateOleObject('Word.Application');
    except
    Wordapp:=GetActiveOleObject('Word.Application');
    end;
    WordApp.Documents.Open(ExtractFilePath(ParamStr(0))+'Shablon.docx');
    WordApp.Selection.Find.ClearFormatting;
    WordApp.Selection.Find.Replacement.ClearFormatting;
    WordApp.Selection.Find.Execute('-FIO-',True,False,False,False,False,True,1,False,ADOQuery1NAME_RU.Value,2);
    WordApp.Selection.Find.Execute('-POL-',True,False,False,False,False,True,1,False,ADOQuery1POL.Value,2);
    WordApp.Selection.Find.Execute('-DOB-',True,False,False,False,False,True,1,False,datetostr(ADOQuery1DOB.Value),2);
    WordApp.Selection.Find.Execute('-ADRS-',True,False,False,False,False,True,1,False,ADOQuery1ADRESS.Value,2);
    WordApp.Selection.Find.Execute('-SPEC-',True,False,False,False,False,True,1,False,ADOQuery1SPEC.Value,2);
    WordApp.ActiveDocument.SaveAs(ExtractFilePath(ParamStr(0))+'WORD_FILES\'+ADOQuery1NAME_RU.Value+'.docx');
    ShellExecute(Handle, 'print', Pchar('WORD_FILES\'+ADOQuery1NAME_RU.Value+'.docx'), nil, nil, SW_HIDE);
    WordApp.Quit;
    WordApp := Unassigned;
    ADOQuery1.Next;
 end;
  end;

procedure TForm4.Button4Click(Sender: TObject);
begin
try
     Excel:=CreateOleObject('Excel.Application');
  except
    Excel:=GetActiveOleObject('Excel.Application');
 end;
   Excel.Workbooks.Open(ExtractFilePath(ParamStr(0))+'Shablon.xlsx');
   FData:=VarArrayCreate([1,ADOQuery1.RecordCount+1,1,ADOQuery1.Fields.Count],varVariant);
   Excel.Worksheets[1].Select;
   for j:=1 to ADOQuery1.Fields.Count do
   FData[1,j]:=ADOQuery1.Fields[j-1].DisplayLabel;
   ADOQuery1.First;
   i:=2;
   while not ADOQuery1.Eof do
      begin
         for j:=1 to ADOQuery1.Fields.Count do
         FData[i,j]:=ADOQuery1.Fields[j-1].Value;
         ADOQuery1.Next;
         i:=i+1;
      end;
   Excel.Range[Excel.Cells[1, 1], Excel.Cells[ADOQuery1.RecordCount+1, ADOQuery1.Fields.Count]].Value:=FData;
   Excel.Range[Excel.Cells[1, 1], Excel.Cells[1,ADOQuery1.Fields.Count]].Select;
   Excel.Selection.Interior.Color:= clSilver;
   Excel.Range[Excel.Cells[1, 1], Excel.Cells[ADOQuery1.RecordCount+1, ADOQuery1.Fields.Count]].Select;
   Excel.Selection.Borders.LineStyle := 1;
   Excel.Selection.Borders.Weight := 2;
   Excel.Columns.AutoFit;
   Excel.Range[Excel.Cells.Item[1, 1], Excel.Cells.Item[1, 1]].Select;
   Excel.visible:=true;
end;

end.
