VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataComboSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents MDS_DCombo As DataCombo
Attribute MDS_DCombo.VB_VarHelpID = -1
Private MDS_ADO As Adodc        'Pointer to the ADO Data Control that the DataCombo is Linked to
Private MDS_RST As Recordset    'Holds a clone of the recordset
Private MDS_FIELD As Field      'Holds the field within the recordset that comprises the ListField of the DataCombo
Private MDS_STOP As Boolean     'This variable determined whether to start a new search or resume the last search
Private MDS_FIELD_NAME As String 'Holds the name of the field specified in MDS_FIELD
'################################################################################################################################################################
'#  Description                                                                                                                                                 #
'#______________________________________________________________________________________________________________________________________________________________#
'#  This class is designed to give the feel of an Access ComboBox to the New DataCombo control in VB6.                                                          #
'#  Whilst this class was designed my Matt Sendt February 2000, you are free to use this Class module or modify it as long as you do not claim it as your       #
'#  intellectual property. This version is only a beta version and will probably require much more error handling, etc.                                         #
'#  Please contact me at sendt@compassnet.com.au if you have any improvements.                                                                                  #
'################################################################################################################################################################
'#  Example usage                                                                                                                                               #
'#______________________________________________________________________________________________________________________________________________________________#
'#  This class is typically used in a form_load event                                                                                                           #
'#  In the declarations section of the form:                                                                                                                    #
'#  Private CC As clsDataComboSearch                                                                                                                            #
'#  In the form_load event:                                                                                                                                     #
'#  Set CC = New clsDataComboSearch                                                                                                                             #
'#  Set CC.Client = DataCombo1                  'Exchange "DataCombo1" for the name of the DataCombo                                                            #
'#  Set CC.ClientRowsource = Me.Adodc2          'NB The ADO DataControl must be the ADODC that contains the recordset for the DataCombo's RowSource property    #
'#                                                                                                                                                              #
'################################################################################################################################################################

Private Sub Class_Initialize()
    Set MDS_DCombo = Nothing
    Set MDS_RST = Nothing
    Set MDS_ADO = Nothing
    Set MDS_FIELD = Nothing
    MDS_FIELD_NAME = ""
End Sub

Private Sub Class_Terminate()
    Set MDS_DCombo = Nothing    'Free memory
    Set MDS_ADO = Nothing
    Set MDS_RST = Nothing
    Set MDS_FIELD = Nothing
End Sub

Public Property Set Client(dcbDCombo As Object)
    If TypeName(dcbDCombo) = "DataCombo" Then   'Check that we are initialising an ADO Datacombo
        Set MDS_DCombo = dcbDCombo
    Else
        Err.Raise Number:=vbObjectError + 1, _
        Source:="clsDataCombo.Client", _
        Description:="The object must be a datacombo."
    End If
End Property

Public Property Set ClientRowsource(adoDatacontrol As Object)
    Dim fld As Field, blnExists As Boolean, strListfield
    If TypeName(adoDatacontrol) = "Adodc" Then
        strListfield = MDS_DCombo.ListField
        Set MDS_ADO = adoDatacontrol
        Set MDS_RST = MDS_ADO.Recordset.Clone(adLockReadOnly)
        With MDS_RST
            blnExists = False
            For Each fld In .Fields
                If fld.Name = strListfield Then
                    MDS_FIELD_NAME = strListfield
                    blnExists = True
                    Exit For
                End If
            Next
        End With
        If Not blnExists Then
            Err.Raise Number:=vbObjectError + 3, _
            Source:="clsDataComboSearch.ClientListField", _
            Description:="The ADO Data Control does not contain the field passed as strListField. The likely cause is that you have passed the incorrect ADODC. Ensure that the ADODC passed is the RowSource of the datacontrol."
        End If
    Else
        Err.Raise Number:=vbObjectError + 2, _
        Source:="clsDataCombo.Client", _
        Description:="The parameter adoDatacontrol must point to an adodatacontrol."
    End If
End Property

Private Sub Highlight(ByVal lngLength As Long)
    With MDS_DCombo
        If Len(.Text) > lngLength Then
            .SelStart = lngLength
            .SelLength = Len(.Text) - lngLength
        End If
    End With
End Sub
'#####################################################################
'#      The following events are inherited from the DataCombo (Sunken) #
'#####################################################################

Private Sub MDS_DCombo_GotFocus()
    On Error GoTo MDS_DCombo_GotFocus_Err
    Set MDS_RST = MDS_ADO.Recordset.Clone(adLockReadOnly)         'Refresh the recordset so that we have the most up-to-date data in the list
    Set MDS_FIELD = MDS_RST.Fields(MDS_FIELD_NAME)
    MDS_STOP = True
MDS_DCombo_GotFocus_Exit:
    Exit Sub
MDS_DCombo_GotFocus_Err:
    Err.Raise Number:=vbObjectError + 1, _
    Source:="clsDataCombo.MDS_DCombo_GotFocus", _
    Description:="There was an error in MDS_DCombo_GotFocus. Likely cause is that the Class has not correctly been initialized."
    Resume MDS_DCombo_GotFocus_Exit
End Sub

Private Sub MDS_DCombo_KeyPress(KeyAscii As Integer)
    On Error GoTo MDS_DCombo_KeyPress_Err
    Static Search As String
    Dim Index As Long
    Dim DoSearch As Boolean
    If MDS_STOP Then Search = ""
    MDS_STOP = False
    Select Case KeyAscii
        Case vbKeyBack
            If Len(Search) Then
                Search = Left(Search, Len(Search) - 1)
                DoSearch = True
            End If
            KeyAscii = 0
        Case vbKeyReturn, vbKeyTab
            With MDS_DCombo
                .SelStart = Len(.Text)
                .SelLength = 0
            End With
            DoSearch = False
            KeyAscii = vbKeyReturn
        Case Is >= vbKeySpace
            Search = Search & UCase(Chr(KeyAscii))
            DoSearch = True
    End Select
    If (DoSearch = True) And Len(Search) > 0 Then
        Dim strCriteria As String, BookMk
        strCriteria = MDS_FIELD.Name & " like '" & Search & "*'"
        With MDS_RST
            BookMk = .Bookmark
            .Find strCriteria, , , 1
            If .EOF Then
                'There is no match
                .Bookmark = BookMk
                Search = Left(Search, Len(Search) - 1)
                If Len(Search) = 0 Then MDS_DCombo.Text = ""
            Else
                MDS_DCombo.Text = MDS_FIELD
            End If
        End With
        Call Highlight(Len(Search))
        KeyAscii = 0
    Else
        If Len(Search) Then
            KeyAscii = KeyAscii
        Else
            MDS_DCombo.Text = ""
        End If
    End If
MDS_DCombo_KeyPress_Exit:
    Exit Sub
MDS_DCombo_KeyPress_Err:
    MsgBox prompt:="Unexpected Error - " & Err.Number & " " & Err.Description, _
    buttons:=vbOKOnly + vbCritical, _
    Title:="clsDataComboSearch.MDS_DCombo_KeyPress"
End Sub

