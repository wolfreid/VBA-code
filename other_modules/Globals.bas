Attribute VB_Name = "Globals"
Public Const keyword As String = "*"
Public Const Destroy As Variant = Null
Public GlobalCancel As Boolean

Public Type Config_Behaviour
    Actions() As Variant ' для таблицы Session
    Profiles() As Variant  'для таблицы Session
    Source_databases() As Variant 'для таблицы Session
End Type

Public Type Config_Lifedata
    Lifestate() As Variant
    Statename() As Variant
End Type

Public Enum AccessLevels
AL_USER = 1
AL_topuser = 2
AL_Admin = 3
AL_DEVELOPER = 9
AL_UNKNOWN = 0
End Enum

Public Enum Sentences
S_Presenting = 1 'без имени
S_Attemption = 2
S_Goodbye = 3
S_Comeback = 4
S_Acceslevel = 5
S_Logout = 6
S_Warning = 7
End Enum

Public Enum Activity
A_on = 1
A_off = 0
End Enum

Public Enum Datasource
DS_life = 0
DS_Session = 1
DS_Projects = 2
DS_PrjStatus = 3
DS_Stages = 4
DS_Calendar = 5
DS_Celebrates = 6
DS_View = 7
End Enum

Public DS As Datasource
Public AT As Activity
Public TitleInterCells As cPrjSysblocks 'К этой переменной не относятся обьекты Userform, прим. для полей
Public Session As cPrjSysblocks 'К этой переменной относятся только обьекты Userform
Public objSetupForm As cConfigurations 'К этой переменной относятся только инстантные обьекты Userform
Public objSetupData As New cConfigurations ' К этой переменной относятся только инстантные обьекты Listobject, существует только один экзмепляр в рамках процедуры
Public CallingAdminButton As String

'Public Type Behaviour_array
'    Quantitatives() As Variant
'    Qualitives() As Variant
'    Auditories() As Variant
'    Stages() As Variant
'    Status() As Variant
'    Responsibilities() As Variant
'     'для таблицы Life
'    Executions() As Variant
'End Type
    
Public Function Behaviour_array() As Config_Behaviour
 With Behaviour_array
  Let .Actions = RangeInArray(range("Action").Columns(2).Cells) 'учитывает все базы
  Let .Source_databases = Get_array_tablesnames 'учитывает все базы
  Let .Profiles = RangeInArray(range("Profile")) 'не учитывает базы
End With
End Function

Public Function Life_array() As Config_Lifedata
 With Life_array
  Let .Lifestate = Array("State in progress", "Complete State", "Wrong")
  Let .Statename = Array("Succesfull", "Incorrect_signing")
End With
End Function

Function Get_EnumSentence(eValue As Sentences, Optional eText As Variant)
  Select Case eValue
    Case 1: Get_EnumSentence = "Добро пожаловать в систему! Ввойдите или обратитесь к админу, чтобы зарегестрироваться."
    Case 2: Get_EnumSentence = "Прежде чем закончить работу в Системе, убедитесь, что Вы сохранили все данные."
    Case 3: Get_EnumSentence = "Всего наилучшего!"
    Case 4: Get_EnumSentence = "C возвращением," & eText & "!"
    Case 5: Get_EnumSentence = eText & ",Ваш уровень доступа:"
    Case 6: Get_EnumSentence = eText & ",подтвердите свой выход ?"
    Case 7: Get_EnumSentence = "Отказ, доступно только зарегестрированному пользователю"
  End Select
End Function

Function Get_EnumLglInforms(eValue As AccessLevels)
  Select Case eValue
  Case 0
    Get_EnumLglInforms = "Пользователь,не прошедший авторизацию в системе, будет иметь всегда этот уровень доступа." & Chr(13) + Chr(10) & _
                        "В распоряжении этому пользователю доступны лист входа и все интерактивности на ней"
  Case 1
    Get_EnumLglInforms = "Авторизированный пользователь будет иметь этот уровень доступа." & Chr(13) + Chr(10) & _
                         "В распоряжении пользователю доступны большинство листов Планировщика и большинство интекрактивностей на них"
  Case 2
    Get_EnumLglInforms = "Авторизированный пользователь с привелегированным статусом будет иметь этот уровень доступа." & Chr(13) + Chr(10) & _
                         "В распоряжении пользователю доступны все листы Планировщика и все интерактивности на них, за исключением баз данных." & _
                         "Добавлена упрощенная навигация."
  Case 3
    Get_EnumLglInforms = "Администратору доступны все страницы, все интерактивности, интерфейс управления аккаунтами"
  Case 9
    Get_EnumLglInforms = "Разработчик - Бог этой системы"
  End Select
End Function



Function Get_ALname(val As AccessLevels) As Integer
If val = AL_DEVELOPER Then Get_ALname = 4 Else Get_ALname = val
End Function





