Version =20
VersionRequired =20
Checksum =-432068803
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8520
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =540
    Top =12
    Right =9180
    Bottom =4332
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x18d05a3b8cc3e240
    End
    GUID = Begin
        0xe3c8312e14757749a9344117ab764a11
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5940
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x1a5e34a07c44ac4d88686db855c464ab
            End
            Begin
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1560
                    Top =300
                    Width =6720
                    Height =5460
                    Name ="Text1"
                    GUID = Begin
                        0xe5f7334b7fbb0348a304c57120fe1b0e
                    End

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =660
                            Top =300
                            Width =780
                            Height =600
                            Name ="Label1"
                            Caption ="Text1:"
                            GUID = Begin
                                0xef04f825ab1c7e4e923b6afb87d72c98
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =120
                    Top =660
                    Width =1320
                    Height =660
                    TabIndex =1
                    Name ="Command1"
                    Caption ="Command1"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0xed609f8cf78c374fa629cf0e7a9baf86
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub AddChildWindows(ByVal hwndParent As Long, ByVal Level As Long)
      
      Dim WT As String
      Dim CN As String
      Dim Length As Long
      Dim hwnd As Long
      
        If Level = 0 Then
          hwnd = hwndParent
        Else
          hwnd = GetWindow(hwndParent, GW_CHILD)
        End If
        Do While hwnd <> 0
          WT = Space(256)
          Length = GetWindowText(hwnd, WT, 255)
          WT = Left$(WT, Length)
          CN = Space(256)
          Length = GetClassName(hwnd, CN, 255)
          CN = Left$(CN, Length)
          Me!Text1 = Me!Text1 & vbCrLf & String(2 * Level, ".") _
                   & WT & " (" & CN & ")"
          AddChildWindows hwnd, Level + 1
          hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop

End Sub

Private Sub Command1_Click()
     
     Dim hwnd As Long
        
        hwnd = GetTopWindow(0)
        If hwnd <> 0 Then
          AddChildWindows hwnd, 0
        End If

End Sub