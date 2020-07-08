Attribute VB_Name = "modAuthor"
'Программа состоит из 4-х глобальных частей:
'1 - обьявление констант
'2 - отображение формы
'3 - реакция на отображение формы
'4 - дополнительные вычисление на выходе реакций

Option Compare Text
'1-обьявление глобальных переменных

Public AL As AccessLevels
Dim s(0 To 255) As Integer, kep(0 To 255) As Integer
Const ACCOUNT_PASSWORD = "jvjh67s23gso@#^%$&^%&(*jkl;kjghc34+"
Const WORKBOOK_ID = "106993", ACCOUNT_INFO_SEPARATOR = "###```###"


Function УровеньДоступа() As String
    УровеньДоступа = AL
End Function


'2.4.1 - проверяем данные ввода из окна ввода
'заводим коллекцию аккаунтов
'если введенній логин соответствует одному из значений логинав коллекции и в том же кортеже паролю -
'-функция возвращает уровень доступа
Function CheckAccount(ByVal Login As String, ByVal Password As String) As AccessLevels
    ' проверяет имя пользователя и пароль среди аккаунтов, сохранённых в файле
    'реализация через функцию с аргументом книга
    ' если такая учётная запись присутствует, функция возвращает уровень доступа
    On Error Resume Next
    arr = AllAccountsArray(ThisWorkbook)
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i, 1)) = UCase(Login) And arr(i, 2) = Password Then
            CheckAccount = arr(i, 3): Exit Function
        End If
    Next i
End Function

'2.4.2 - аккаунты вызываются из коллекции. Коллекция образается через функцию чтения аккаунтв в книге
'вызов происходит дважды: при вызове формы и событии входа. Процедура загружает действующие аккаунты,
'потом проверка на наличие действующих аккаунтов при отсутствии функция завершается, все наследующие эту функцию
'завершаются
'Коллекция собирает все аккаунты, один аккаунт - один item.
'задается двумерній массив с колличеством аккаунтов(коллекция начинается с нуля поєтому -1 и 4 столбцами подобной информации)
'Цикл для каждого єлемента коллекции, тоесть аккаунта, записи коллекции можно сказать выполняем следующие операции
'кодировался пароль и заполнитель с уровнем доступа
'
Function AllAccountsArray(ByRef Wb As Workbook) As Variant
    ' возвращает двумерный массив размерностью Кол-воАккаунтов * 4
    '4 колонки двумерного массива отводится под: 0-индекс,1-логин,2-пароль,3-уровень доступа
    Dim coll As Collection: Set coll = ReadAllAccounts(Wb) '1
    If coll.Count = 0 Then Exit Function
    ReDim arr(0 To coll.Count - 1, 0 To 3): On Error Resume Next
    For i = 1 To coll.Count
        arrTEMP = Split(coll(i), ACCOUNT_INFO_SEPARATOR) 'элемент коллекции разбивается на подстроки с помощью имеющегося разделителя
        txt = EnDeCrypt(arrTEMP(1), ACCOUNT_PASSWORD) 'дешефрация пароля'2, аргументами является шифр и ключ дешифровки
        arr(i - 1, 0) = arrTEMP(2)    ' 0 - индекс
        arr(i - 1, 1) = arrTEMP(0)    ' 1 - логин
        arr(i - 1, 2) = Split(txt, ACCOUNT_INFO_SEPARATOR)(2)    ' 2 - пароль
        arr(i - 1, 3) = val(Split(txt, ACCOUNT_INFO_SEPARATOR)(1))    ' 3 - уровень доступа
    Next i
    AllAccountsArray = arr
End Function

'2.4.2.1 - коллекция собирает все упорядоченные записи в свойстве книги CustomDocumentProperties
'индекс извлекается с помощью строковой функции и оператора сравнивания с шаблоном, который
'уже заранее был прописан как константа для индексации в Свойство книги
' индекс принадлежит к номерным значениям и закреплен за аккаунтом: именем и паролем
'если получаемое значение из функции mid как число значит можно через него обратится за аккаунтом
'информация об аккаунте компануется в текстовой переменной acc,
'cкелет выражения имеет такое строение: ЛОГИН+разделитель+ПАРОЛЬВШИФРЕ+разделитель+ИНДЕКС
Function ReadAllAccounts(ByRef Wb As Workbook) As Collection
    Set ReadAllAccounts = New Collection: Dim acc As String, ind As String
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name Like "Login#*" Then
                ind = Mid(cdp.name, 6) 'на віходе число
                If ind Like String(Len(ind), "#") Then 'на выходе
                    AccountInfo = GDoc(Wb, "AccountInfo" & ind) '1
                    acc = cdp.Value & ACCOUNT_INFO_SEPARATOR & AccountInfo & ACCOUNT_INFO_SEPARATOR & val(ind)
                    ReadAllAccounts.Add acc, acc
                End If
            End If
        Next
    End If
End Function

'2.0.1 - вызов функции в отображаемой форме, которая предостовляет доступ к настройкам проекта,
'при первом запуске key нулевое - информации о записи нет, результатом функции тогда будет True, в последующих
'запусках функция возвращает в форму False

Function FirstRun() As Boolean
    FirstRun = GetSetting(Application.name, "Authentification_Gant", WORKBOOK_ID, "") = ""
End Function
'2.4.2.1 - обращение происходило через координаты книги и ключевого слова "AccountInfo" с индексом,
' в этой функции берется пароль аккаунта, accauntinfo - собственное имя записи, идентификатор условно
Function GDoc(ByRef Wb As Workbook, ByVal VarName As String) As String
    ' чтение переменной из книги Excel
    ' функция возвращает значение пользовательского свойства VarName
    ' (если нужное пользовательское свойство отсутствует, возвращает пустую строку)
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name = VarName Then GDoc = cdp.Value
        Next
    End If
End Function

'2.4.2 - аккаунты проходят дешифровку по каждому значению из ascii кодировки
'для каждого значения a в ключе отбирается одно значение согласно его длине, цикл на 256 проходов,
'для a больше длины ключа b = 1, уникальных числовых результатов в количестве длины ключа
'для всех остальніх значений в порядке превышабщих длину ключа kep равен ascii первого значения
' по строчку 11 подготовка ключа рассшифровки
'c 12 рассшифровка пароля
Public Function EnDeCrypt(ByVal plaintxt As String, ByVal Password As String) As String
    Dim Temp As Integer, a As Integer, b As Integer, cipherby As Byte, cipher As String
    b = 0
    For a = 0 To 255
        b = b + 1
        If b > Len(Password) Then b = 1
        kep(a) = Asc(Mid$(Password, b, 1))
    Next a '
    For a = 0 To 255: s(a) = a: Next a: b = 0 'собирается массив уник значений от 256 штук
'в 256 проходах суммируются числа, на старте (b=1+a=0+kep = asc(j))%256. Результат - остатки
'переменная b получает последовательность из остатков вместо простого порядка в пред. цикле
'переменная temp присваивает s(a)
'элемент массива заменяется на элемент этого массива по индексу в качестве остатка b
'в то же время элемент массива по остатку меняется на тот что был ранее в s(a)
    For a = 0 To 255: b = (b + s(a) + kep(a)) Mod 256: Temp = s(a): s(a) = s(b): s(b) = Temp: Next a
'когда все элементы обращены согласно функции mod и переставлены согласно выражению подмены происходит следующее
'на выходе новая последовательность элементов с 256 уникальными значениями от 0 до 255,
'собственно рассшифровка пароля, задан цикл по длине зашифрованного пароля. в каждом проходе цикла совершается
'дешифровка определенного символа. Аргументы: байтовое выражение символа
'полученное правильное значение от 0 до 255
    For a = 1 To Len(plaintxt): cipherby = EnDeCryptSingle(Asc(Mid$(plaintxt, a, 1)))
        cipher = cipher & Chr(cipherby): Next: EnDeCrypt = cipher
End Function

'в дешифровке символа задйствуются функция где задана постоянная точка отсчета 1 mod 256 = 1
'в дешифровке символа задйствуются функция где задана постоянная точка отсчета j = s(1) - первое значение последовательности s()
'повторна перестановка, в которой подменяется первое значение последовательности s(i) на значение в s(j), а s(j) на s(i)
'целевая функция, получние k от деления символа в последовательности по индексу из сум первого значения и j-го значения
'в сл. выражении ascii значение символа сравнивается с k значением побитово. Результатом является значение в байте
Public Function EnDeCryptSingle(plainbyte As Byte) As Byte
    Dim i As Integer, j As Integer, Temp As Integer, k As Integer, cipherby As Byte
    i = (i + 1) Mod 256: j = (j + s(i)) Mod 256: Temp = s(i): s(i) = s(j): s(j) = Temp
    k = s((s(i) + s(j)) Mod 256): cipherby = plainbyte Xor k: EnDeCryptSingle = cipherby
End Function

'2.5.1.1 сортируется список аккаунтов
'двумерные массивы состоят из строк и столбцов. Первым измерением считается количество строк, вторым - количество столбцов.
'двумерный массив array(x,y)
'x-для последнего значения в столбце испльзуется аргумент 1 2-мерного массива. в матрице - позиция внизу столбца
'y -для последнего значения  в строке используется аргумент 2 2-мерного массива
'Значение массива = 3 - 4 столбца
'Do until работает как счетчик строк
Public Function CoolSort(SourceArr As Variant) As Variant
    ' сортировка двумерного массива по нулевому столбцу
    Dim Check As Boolean, iCount As Integer, jCount As Integer, nCount As Integer
    ReDim tmpArr(UBound(SourceArr, 2)) As Variant
    Do Until Check
        Check = True
        For iCount = LBound(SourceArr, 1) To UBound(SourceArr, 1) - 1 'от первого до последнего значения в столбце
            If val(SourceArr(iCount, 0)) > val(SourceArr(iCount + 1, 0)) Then 'отдельные значения в первом столбце
                For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2) 'найденная строка с большим значением
                    tmpArr(jCount) = SourceArr(iCount, jCount) 'во временном трехзначимым массиве скл. все значения аккаунта
                    SourceArr(iCount, jCount) = SourceArr(iCount + 1, jCount)
                    SourceArr(iCount + 1, jCount) = tmpArr(jCount)
                    Check = False
                Next
            End If
        Next
    Loop
    CoolSort = SourceArr
End Function

'В delete account удаляется одиночный документ аккаунта: запись о логине и запись о пароле
'аргументами служит первая колонка - идентификатор строки базы

Sub DeleteAccount(ByVal index As Long)
    On Error Resume Next
    DDoc ThisWorkbook, "Login" & index
    DDoc ThisWorkbook, "AccountInfo" & index
End Sub

Sub DDoc(ByRef Wb As Workbook, ByVal VarName As String)
    ' удаление пользовательского свойства из книги Excel
    If Wb.CustomDocumentProperties.Count > 0 Then    ' если они вообще есть
        For Each cdp In Wb.CustomDocumentProperties    ' перебираем все свойства
            If cdp.name = VarName Then cdp.Delete: Exit Sub    ' удаляем
        Next
    End If
End Sub

' по сути обработчик новых пользователей
'имеет сборщик SDoc для логина, для пароля с аргументами: идентификатор с расспознователем "Login","Accountinfo"
'пароль зашифровать заранее установив сепарацию ключевых свойств
'зашифровка пароля происходит здесь, для шифрования задается обратный порядок
Sub AddAccount(ByVal Login As String, ByVal Password As String, _
               ByVal AccessLevel As AccessLevels, ByVal index As Long)
    ' добавляет учётную запись в файл
    SDoc ThisWorkbook, "Login" & index, Login
    AccountInfo = ACCOUNT_INFO_SEPARATOR & Format(AccessLevel, "0000") & ACCOUNT_INFO_SEPARATOR & Password
    SDoc ThisWorkbook, "AccountInfo" & index, EnDeCrypt(AccountInfo, ACCOUNT_PASSWORD)
    SaveSetting Application.name, "Authentification_Gant", WORKBOOK_ID, "Accounts created"
End Sub

'В ходе добавления нового пользователя, проверка на наличии имеющегося. Если есть - удалить, но лучше
'свернуть секцию по созданию нового пользователя с обнулением всех текущих изменений
'Занести информацию в новую строку, повторить для AccessInfo
Sub SDoc(ByRef Wb As Workbook, ByVal VarName As String, ByVal VarValue As Variant)
    ' сохранение пользовательского свойства в книге Excel
    DDoc Wb, VarName    ' удаляем свойство, если оно уже есть
    ' и создаём новое с нужным значением
    Wb.CustomDocumentProperties.Add VarName, False, msoPropertyTypeString, CStr(VarValue)
End Sub
'Функция по типу процедуры, выполняется после нажатия "Удалить все листы"
'Цикл проходит по всем записям и удаляет их методом Delete,
'В конце запись сохраняется в реестр, в дальнейшем можно развить концепцию до легитимизации доступа на компьютере
Function DeleteAllAccounts()
    Dim Wb As Workbook: Set Wb = ThisWorkbook
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name Like "Login#*" Or cdp.name Like "AccountInfo#*" Then
                cdp.Delete
            End If
        Next
    End If
    SaveSetting Application.name, "Authentification_Gant", WORKBOOK_ID, ""
End Function

Private Sub Test_CreateAccounts()
    AddAccount "admin", "admin", AL_Admin, 1
    AddAccount "user1", "password", AL_USER, 2
    AddAccount "Developer", "1", AL_DEVELOPER, 0
End Sub
