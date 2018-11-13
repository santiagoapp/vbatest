VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TipoDeMaquinaMainForm 
   Caption         =   "Administraci�n de los tipos de maquinas"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   OleObjectBlob   =   "TipoDeMaquinaMainForm.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TipoDeMaquinaMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TipoDeMaquina As cTipoDeMaquina
Private PreguntaCheckList As cPreguntaCheckList
Private Form As cForm

Private Sub agregar_Click()

    Set AgregarForm = UserForms.Add("TipoDeMaquinaAddForm")
    AgregarForm.show
    Unload AgregarForm
    Set AgregarForm = Nothing
    Call Me.getFields
    
End Sub

Private Sub borrar_Click()
    
    mensaje = MsgBox("�Desea continuar con la eliminaci�n del registro?", vbInformation + vbYesNo)
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        If mensaje = vbYes Then
            response = TipoDeMaquina.delete(CInt(Me.ListBox1))
            If response Then MsgBox "Registro eliminado con �xito", vbInformation
        End If
    End If
    Call Me.getFields
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub modificar_Click()
    
    If VarType(Me.ListBox1) = vbNull Then
        MsgBox "Por favor seleccione un registro para continuar", vbInformation, "Seleccionar registro"
    Else
        Set EditarForm = UserForms.Add("TipoDeMaquinaAddForm")
        Set TipoDeMaquina = New cTipoDeMaquina
        EditarForm.ID = Me.ListBox1
        arr = TipoDeMaquina.show( _
            colsFilter:=Array("id"), _
            logicOperators:=Array("="), _
            colsValues:=Array(CInt(Me.ListBox1)) _
        )
        EditarForm.Caption = "Editar registro"
        EditarForm.Frame1.Caption = "Editar tipo de m�quina"
        
        EditarForm.Tipo = arr(1, 0)

        EditarForm.show
        Unload EditarForm
        Set EditarForm = Nothing
    End If
    Call Me.getFields
    
End Sub

Private Sub preguntasBtn_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub preguntasBtn_Click()
    
    Set AsignarForm = UserForms.Add("AsignarPreguntasForm")
    Set PreguntaCheckList = New cPreguntaCheckList
    AsignarForm.tipoDeMaquinaID = CInt(Me.ListBox1)
    AsignarForm.TipoField = Me.ListBox1.List(, 1)
    
    Call AsignarForm.getFields
    AsignarForm.show
    Unload AsignarForm
    Set AsignarForm = Nothing
    
End Sub

Private Sub UserForm_Initialize()

    Set TipoDeMaquina = New cTipoDeMaquina
    Set Form = New cForm
    Call Me.getFields
    
End Sub
Public Function getFields()

    arr = TipoDeMaquina.show( _
        fields:=Array("id", "tipo_maquina") _
    )
    Call Form.fillListOrComboBox(arr, Me.ListBox1)
    
End Function






