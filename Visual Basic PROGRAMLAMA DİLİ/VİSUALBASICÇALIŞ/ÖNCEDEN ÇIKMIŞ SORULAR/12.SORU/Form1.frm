VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fibonocci Dizisi"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   1800
      TabIndex        =   2
      Top             =   5280
      Width           =   10695
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hesapla"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, z, sayi As Integer

Sub fibonacciform(sayi) ''procedure
x = 0
y = 1
Print (" ")
''Print (x & " ")
Print (y)



For i = 0 To sayi - 2 Step 1
z = x + y
Print (" " & z)

x = y
y = z
Next i

End Sub

Sub fibonaccitext(sayi) ''procedure
x = 0
y = 1
Text1.Text = (" ")
''text1.text = (x & " ")
Text1.Text = (y)


For i = 0 To sayi - 2 Step 1
z = x + y
Text1.Text = Text1.Text & (" " & z)

x = y
y = z
Next i

End Sub

Sub fibonacciliste(sayi) ''procedure
x = 0
y = 1
List1.Clear
''List1.AddItem (x & " ")
List1.AddItem (y & " ")


For i = 0 To sayi - 2 Step 1
z = x + y
List1.AddItem (z)
x = y
y = z
Next i

End Sub


Private Sub Command1_Click()
sayi = InputBox("Lütfen Adým Sayýsýný Girin", "Adým Sayýsý")
fibonacciliste (sayi)
fibonaccitext (sayi)
fibonacciform (sayi)


End Sub


''int x = 0;
  ''          int y = 1, z;
    ''        int sayi;
''
  ''          Console.WriteLine("Adým sayýsýný giriniz:");
    ''        sayi = Convert.ToInt32(Console.ReadLine());
      ''      Console.WriteLine();
''
  ''          Console.WriteLine(x + " " + "\n" + y + " ");

    ''        for(int i=0; i< sayi-2; i++)
      ''      {
        ''        z = x + y;
          ''      Console.WriteLine( z );
            ''    x = y;
              ''  y = z;

            ''}
