VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCorrectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'
'                                   1. VARIABLES
'
'------------------------------------------------------------------------------------'
Private tabla As String
Private columns As Variant
Private values As Variant

Private pID As Variant
Private pEquipoID As Variant
Private pEncargadoID As Variant
Private pTipoDeFalla As Variant
Private pCausaDeFalla As Variant
Private pProvedor As Variant
Private pCosto As Variant
Private pFechaInicio As Variant
Private pFechaFin As Variant
Private pCreatedAt As Variant
Private pUpdatedAt As Variant
Private pDeletedAt As Variant

Private softDeletes As Boolean
Private Con As cConnection
Private QueryBuilder As cQueryBuilder

'------------------------------------------------------------------------------------'
'
'                                   2. PROPIEDADES
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   2.1 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get ID() As Variant
    ID = pID
End Property

Public Property Let ID(value As Variant)
    pID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.2 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get equipoID() As Variant
    equipoID = pEquipoID
End Property

Public Property Let equipoID(value As Variant)
    pEquipoID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get encargadoID() As Variant
    encargadoID = pEncargadoID
End Property

Public Property Let encargadoID(value As Variant)
    pEncargadoID = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get tipoDeFalla() As Variant
    tipoDeFalla = pTipoDeFalla
End Property

Public Property Let tipoDeFalla(value As Variant)
    pTipoDeFalla = value
End Property


'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get causaDeFalla() As Variant
    causaDeFalla = pCausaDeFalla
End Property

Public Property Let causaDeFalla(value As Variant)
    pCausaDeFalla = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get provedor() As Variant
    provedor = pProvedor
End Property

Public Property Let provedor(value As Variant)
    pProvedor = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get Costo() As Variant
    Costo = pCosto
End Property

Public Property Let Costo(value As Variant)
    pCosto = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get fechaInicio() As Variant
    fechaInicio = pFechaInicio
End Property

Public Property Let fechaInicio(value As Variant)
    pFechaInicio = value
End Property

'------------------------------------------------------------------------------------'
'
'   2.3 PROPIEDAD
'
'------------------------------------------------------------------------------------'

Public Property Get fechaFin() As Variant
    fechaFin = pFechaFin
End Property

Public Property Let fechaFin(value As Variant)
    pFechaFin = value
End Property

'------------------------------------------------------------------------------------'
'
'                                   3. CONSTRUCTOR
'
'------------------------------------------------------------------------------------'

Private Sub Class_Initialize()
    
    Set Con = New cConnection
    Set QueryBuilder = New cQueryBuilder
    
    softDeletes = True
    Con.DBPath = ThisWorkbook.Path & "\db\db.accdb"
    columns = Array("id", "equipo_id", "encargado_id", "tipo_de_falla", "causa_de_falla", "provedor", "costo", "fecha_inicio", "fecha_fin", "created_at", "updated_at", "deleted_at")
    tabla = "correctivos"
    QueryBuilder.tabla = tabla
    QueryBuilder.softDeletes = softDeletes
    
End Sub

'------------------------------------------------------------------------------------'
'
'                                   4. M�TODOS
'
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
'
'   4.1 CREAR NUEVO
'
'------------------------------------------------------------------------------------'

Public Function create() As Boolean
    
    pCreatedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    pUpdatedAt = pCreatedAt
    values = Array(pID, pEquipoID, pEncargadoID, pTipoDeFalla, pCausaDeFalla, pProvedor, pCosto, pFechaInicio, pFechaFinpCreatedAt, pUpdatedAt, pDeletedAt)
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.insertQry
    If Con.setRecords Then create = True Else create = False
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 MOSTRAR
'
'------------------------------------------------------------------------------------'

Public Function show( _
    Optional fields As Variant = Empty, _
    Optional colsFilter As Variant = Empty, _
    Optional logicOperators As Variant = Empty, _
    Optional colsValues As Variant = Empty, _
    Optional joins As Variant = Empty _
) As Variant
    
    QueryBuilder.columns = fields
    QueryBuilder.colsFilters = colsFilter
    QueryBuilder.logicOperators = logicOperators
    QueryBuilder.colsValues = colsValues
    QueryBuilder.joins = joins
    
    Con.qry = QueryBuilder.showQry
    show = Con.getRecords
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 ACTUALIZAR
'
'------------------------------------------------------------------------------------'

Public Function update(key As String) As Boolean
    
    pUpdatedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    values = Array(pEquipoID, pEncargadoID, pTipoDeFalla, pCausaDeFalla, pProvedor, pCosto, pFechaInicio, pFechaFinpCreatedAt, pUpdatedAt, pDeletedAt)
    columns = Array("equipo_id", "encargado_id", "tipo_de_falla", "causa_de_falla", "provedor", "costo", "fecha_inicio", "fecha_fin", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.updateQry("id", key)
    update = Con.setRecords
    
End Function

'------------------------------------------------------------------------------------'
'
'   4.2 ELIMINAR
'
'------------------------------------------------------------------------------------'

Public Function delete(key As String) As Boolean
    
    pDeletedAt = Application.Text(Now, "dd/mm/yyyy H:MM")
    values = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, pDeletedAt)
    columns = Array("id", "equipo_id", "encargado_id", "tipo_de_falla", "causa_de_falla", "provedor", "costo", "fecha_inicio", "fecha_fin", "created_at", "updated_at", "deleted_at")
    
    QueryBuilder.columns = columns
    QueryBuilder.values = values
    
    Con.qry = QueryBuilder.deleteQry("id", key)
    delete = Con.setRecords
    
End Function








