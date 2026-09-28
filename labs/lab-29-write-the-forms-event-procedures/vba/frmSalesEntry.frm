VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSalesEntry 
   Caption         =   "NorthWind Sales Entry"
   ClientHeight    =   4200
   ClientWidth     =   6000
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblRepID
      Caption         =   "Rep ID"
      Height          =   24
      Left            =   12
      TabIndex        =   0
      Top             =   12
      Width           =   96
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtRepID
      Caption         =   ""
      Height          =   24
      Left            =   108
      TabIndex        =   1
      Top             =   12
      Width           =   96
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblRepName
      Caption         =   "Rep Name"
      Height          =   24
      Left            =   12
      TabIndex        =   2
      Top             =   42
      Width           =   96
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtRepName
      Caption         =   ""
      Height          =   24
      Left            =   108
      TabIndex        =   3
      Top             =   42
      Width           =   96
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblRegion
      Caption         =   "Region"
      Height          =   24
      Left            =   12
      TabIndex        =   4
      Top             =   72
      Width           =   96
   End
   Begin {8BD21D30-EC42-11CE-9E0D-00AA006002F3} cboRegion
      Caption         =   ""
      Height          =   24
      Left            =   108
      TabIndex        =   5
      Top             =   72
      Width           =   96
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblSalesAmount
      Caption         =   "Sales Amount"
      Height          =   24
      Left            =   12
      TabIndex        =   6
      Top             =   102
      Width           =   96
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtSalesAmount
      Caption         =   ""
      Height          =   24
      Left            =   108
      TabIndex        =   7
      Top             =   102
      Width           =   96
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdSubmit
      Caption         =   "Submit"
      Height          =   24
      Left            =   12
      TabIndex        =   8
      Top             =   132
      Width           =   84
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdCancel
      Caption         =   "Cancel"
      Height          =   24
      Left            =   12
      TabIndex        =   9
      Top             =   132
      Width           =   84
   End
End
Attribute VB_Name = "frmSalesEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    ' Clear the inputs and load the region list as the form opens
    txtRepID.Value = ""
    txtRepName.Value = ""
    txtSalesAmount.Value = ""

    cboRegion.Clear
    cboRegion.AddItem "Central"
    cboRegion.AddItem "East"
    cboRegion.AddItem "North"
    cboRegion.AddItem "West"

End Sub


Private Sub cmdSubmit_Click()

    Dim lngRow As Long

    Worksheets("Sales").Activate

    ' first empty row below the existing data
    lngRow = Worksheets("Sales").Range("A1048576").End(xlUp).Row + 1

    Cells(lngRow, 1) = txtRepID.Value
    Cells(lngRow, 2) = txtRepName.Value
    Cells(lngRow, 3) = cboRegion.Value
    Cells(lngRow, 5) = txtSalesAmount.Value

    ' reset the form for the next entry
    txtRepID.Value = ""
    txtRepName.Value = ""
    txtSalesAmount.Value = ""
    cboRegion.Value = ""

    MsgBox "Record added to row " & lngRow & "."

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
