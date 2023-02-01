VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ws As Worksheet
Public id As Variant
Public triangle As Collection
Public ABC As Collection
Public nbAutre As Collection
Public accord As Collection
Public mission As Collection
Public ressortissants As Collection
Public nom As Collection
Public indiceEco As Collection
Public indiceAutre As Collection
Public indiceSecu As Collection
Public indiceDiplo As Collection
Public indiceCult As Collection
Public indiceAll As Collection
Public TRUC As Collection
Public cvert As Collection
Public cbleu As Collection
Public UE As Collection
Public cviolet As Collection
Public TextPays As Collection
Public TextAutre As Collection
Public OpAutre As Collection
Public rg As Range

'objet qui permet d'avoir accès à toutes les propriétés d'un pays qui se trouve dans le tableau "Synthèse affichage"

Public Sub init()
    Dim i As Integer
    Set Me.ws = ThisWorkbook.Sheets("SynthèseAffichage")
    Set Me.triangle = New Collection
    Set Me.ABC = New Collection
    Set Me.nbAutre = New Collection
    Set Me.TRUC = New Collection
    Set Me.cvert = New Collection
    Set Me.UE = New Collection
    Set Me.cbleu = New Collection
    Set Me.nom = New Collection
    Set Me.indiceEco = New Collection
    Set Me.indiceAutre = New Collection
    Set Me.indiceSecu = New Collection
    Set Me.indiceDiplo = New Collection
    Set Me.indiceCult = New Collection
    Set Me.indiceAll = New Collection
    Set Me.cviolet = New Collection
    Set Me.TextPays = New Collection
    Set Me.TextAutre = New Collection
    Set Me.OpAutre = New Collection
    Set rg = Me.ws.ListObjects("TD_Synthese").DataBodyRange
    Me.id = Me.ws.ListObjects("TD_Synthese").ListColumns("Code excel").DataBodyRange.Value
    For i = 1 To rg.Rows.Count
        Me.nom.Add Item:=rg(i, 1).Value, Key:=CStr(rg(i, 2).Value)
        Me.indiceAutre.Add Item:=rg(i, 6).Value, Key:=CStr(rg(i, 2).Value)
        Me.indiceAll.Add Item:=rg(i, 9).Value, Key:=CStr(rg(i, 2).Value)
        Me.triangle.Add Item:=rg(i, 10).Value, Key:=CStr(rg(i, 2).Value)
        Me.nbAutre.Add Item:=rg(i, 11).Value, Key:=CStr(rg(i, 2).Value)
        Me.OpAutre.Add Item:=rg(i, 12).Value, Key:=CStr(rg(i, 2).Value)
        Me.TRUC.Add Item:=rg(i, 13).Value, Key:=CStr(rg(i, 2).Value)
        Me.cvert.Add Item:=rg(i, 14).Value, Key:=CStr(rg(i, 2).Value)
        Me.cbleu.Add Item:=rg(i, 15).Value, Key:=CStr(rg(i, 2).Value)
        Me.UE.Add Item:=rg(i, 16).Value, Key:=CStr(rg(i, 2).Value)
        Me.cviolet.Add Item:=rg(i, 17).Value, Key:=CStr(rg(i, 2).Value)
        Me.TextPays.Add Item:=rg(i, 19).Value, Key:=CStr(rg(i, 2).Value)
        Me.TextAutre.Add Item:=rg(i, 20).Value, Key:=CStr(rg(i, 2).Value)
    Next
End Sub

