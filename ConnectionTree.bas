Attribute VB_Name = "ConnectionTree"
Public Fathers As New Dictionary
Public Branches As New Dictionary
Public Probabilities As New Dictionary
Public ShapeT As New Dictionary



Sub ConnsDic(Snode As Shape, Enode As Shape)


Dim ws As Worksheet
Dim ar As New Collection
Set ws = Worksheets("Structuring")

If Snode.AutoShapeType = msoShapeFlowchartDecision Then
            'Add connection to Father
        If Fathers.Exists(Snode.Name) Then 'If a branch has been added for this decision
                Fathers(Snode.Name).Add Enode.Name 'add the new son
        Else 'if it is the first branch
                ar.Add Enode.Name
                Fathers.Add Snode.Name, ar
        End If
       If Not Branches.Exists(Enode.Name) Then 'if there is no branch already created starting in Enode
            Branches.Add Enode.Name, "" 'start a new branch
        End If

    
Else
    Fathers.Add Snode.Name, Enode.Name 'if the start node is an activity
    If Snode.Name = "1" Then
        Branches.Add Snode.Name, "" 'if it is the first node, start a branch
        ShapeT.Add Snode.TextFrame.Characters.text, Snode.AutoShapeType 'save the type of shape it is
        If Not ShapeT.Exists(Enode.Name) Then
            ShapeT.Add Enode.TextFrame.Characters.text, Enode.AutoShapeType 'save the type of shape it is
        End If
        Exit Sub
    End If

End If

If Not ShapeT.Exists(Snode.TextFrame.Characters.text) Then
  ShapeT.Add Snode.TextFrame.Characters.text, Snode.AutoShapeType 'save the type of shape it is
End If

If Not ShapeT.Exists(Enode.TextFrame.Characters.text) Then
      ShapeT.Add Enode.TextFrame.Characters.text, Enode.AutoShapeType 'save the type of shape it is
End If
End Sub


