Attribute VB_Name = "modulegeneral"
Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Public customersformisopen As Boolean
Public suppliersformisopen As Boolean
Public titlesformisopen As Boolean
Public titleaddmodformisopen As Boolean
Public purchasesformisopen As Boolean
Public salesformisopen As Boolean
Public user As String
Public role As String
Public enable As Boolean

Public Sub FillCombo(srscombo As ComboBox, srsarr As Variant, srsdata As Variant)

'This procedure fills a combo box with field name from a given Array
'used by the search form for Searching
    For i = LBound(srsarr) To UBound(srsarr)
        srscombo.AddItem srsarr(i)
        srscombo.ItemData(i) = srsdata(i)
    Next i
End Sub

Sub Main()
    If App.PrevInstance = True Then
        MsgBox "Application is already running", vbInformation
        End
    End If
    Dim constr As String
    constr = "Data Source=" & App.Path & "\Database\inventory.mdb;Persist Security Info=False;Jet OLEDB:Database Password=thatstherightway"
    Data1.conn.Open constr
    Dim rs As New ADODB.Recordset
    rs.Open "select * from users", Data1.conn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        rs.Close
        Load frmSplash
        maketransparent frmSplash
        frmSplash.Show vbModal
        frmadminpass.Show vbModal
    Else
        rs.Close
        Load frmSplash
        maketransparent frmSplash
        frmSplash.Show vbModal
        DoEvents
        frmlogin.Show vbModal
    End If
End Sub
Public Sub maketransparent(frm As Form)
    On Error GoTo err:
    SetWindowLongA frm.hwnd, GWL_EXSTYLE, GetWindowLongA(frm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hwnd, 0, 200, LWA_ALPHA
err:
    Exit Sub
End Sub

