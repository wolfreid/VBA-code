VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public activeForm As Integer
Private mformid As String


Public Property Let GetNameForm(rVal As String)
 mformid = rVal
End Property

Public Property Get KeepIdForm() As Integer
Select Case mformid
    Case "Введите имя пользователя или пароль": KeepIdForm = 0
    Case "Leaving": KeepIdForm = 1
    Case "Creating": KeepIdForm = 2
End Select
End Property

Public Property Get Database(native As Datasource) As String
'Здесь устанавливается значение для каждой базы, которое не привязано к истинному имени таблицы - для базы данных
Select Case native
    Case 0: Database = "Life"
    Case 1: Database = "Session"
    Case 2: Database = "Projects"
    Case 3: Database = "PrjStatus"
    Case 4: Database = "Stages"
    Case 5: Database = "Calendar"
    Case 6: Database = "Celebrates"
    Case 7: Database = "View"
End Select
End Property



