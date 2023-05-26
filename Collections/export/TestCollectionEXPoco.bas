Attribute VB_Name = "TestCollectionEXPoco"
'@Folder("VBAProject")
Option Explicit

Public Sub DoTestPoco()
    Dim PocoColl As Collection
    Set PocoColl = New Collection
    PocoColl.Add Item:=TestPOCO.Create("Gamma"), Key:="Gamma"
    PocoColl.Add Item:=TestPOCO.Create("Omega"), Key:="Omega"
    PocoColl.Add Item:=TestPOCO.Create("Zeta"), Key:="Zeta"
    
    Dim v As Variant
    v = CStr("Not a POCO")
    PocoColl.Add Item:=v, Key:="NotAPoco"
    
    Debug.Assert CollectionEx.From(PocoColl).ContainsByProperty("Name", "NonexistingPropValue") = False
    Debug.Assert CollectionEx.From(PocoColl).ContainsByProperty("Name", "Gamma") = True
    Debug.Print "ContainsByProperty BadPropName = "; CollectionEx.From(PocoColl).ContainsByProperty("BadPropName", "Gamma")

    ' TestPOCO in the next line is the default instantiation because it has attribute PredeclareId
    ' It will also check TypeOf on each item in the collection
    CollectionEx.From(PocoColl).ForEach TestPOCO, "HandlePOCO"
End Sub


