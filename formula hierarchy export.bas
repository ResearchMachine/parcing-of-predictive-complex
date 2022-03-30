Attribute VB_Name = "Module1"
Option Explicit



'Application.SendKeys "^g ^a {DEL}"

Sub Macro()
    Open ThisWorkbook.Path & "\template.txt" For Output As #1
    'Print #1, "Hello"

    'Dim strKolvoAndSpisokFormulCharov As String
    Dim strHelp As String 'Строковый Помогалка
    Dim strHelp2 As String 'Строковый Помогалка
    Dim lngHelp As Long 'Целочисленный Помогалка
    Dim lngHelp2 As Long 'Целочисленный Помогалка
    Dim strCurrentList As String 'Текущий лист
    Dim strCurrentChar As String 'Текущая ячейка
    Dim lngIterChar As Long 'Целочисленный итератор по таблице
    Dim strIter As Variant 'Строковый итератор
    Dim strMainFormula As String 'Формула
    Dim strIterChar As String 'Переменная для подстановки ячейки
    Dim arrSearchField(1 To 2, 1 To 1) As Variant 'Область поиска
    Dim lngHowMuchForm As Long 'Счетчик формул
    Dim lngWhereIterCharInFormula As Long 'Где в формуле ячейка
    Dim strCurrentFormula As String 'Текущая формула на обработке
    Dim lngWhereListIterCharInFormula As Long 'Константа для поиска в листа в формуле
    Dim strListIterChar As String 'Текущая формула на обработке
    'Dim strKolvoAndSpisokFormulCharov As String 'Выходное
    Dim strType As String 'Тип переменной формула или число
    Dim strLeftNomerForm As String 'Строковый Помогалка
    Dim strCurrentKolvoAndSpisokFormulCharov As String 'Строковый Помогалка
    Dim lngNYarus As Long 'Номер яруса
    Dim lngNForm As Long 'Номер формулы
    Dim strVivodPredYarus As String 'Вывод предыдущего яруса
    Dim strVivodTekYarus As String 'Вывод текущего яруса
    Dim strObshSpisokForm As String 'Общий список формул ОбСпФо
    Dim lngIter As Long 'Целочисленный итератор
    Dim lngKolvoFormulPredYarus As Long 'Колво формул на предыдущем ярусе
    Dim lngKolvoFormulTekYarus As Long 'Колво формул в текущем ярусе
    
    strCurrentList = "Расчет дебитов"
    strCurrentChar = "B2"
    strLeftNomerForm = ""
    lngNYarus = 1
    strVivodPredYarus = strKolvoAndSpisokFormulCharov(strCurrentList, strCurrentChar, lngNYarus) 'Вывод предыдущего яруса
    lngNYarus = 1 'Номер яруса
    lngNForm = 1 'Номер формулы
    strVivodTekYarus = "" 'Вывод текущего яруса
    
    Print #1, strVivodPredYarus + strCurrentList
    Debug.Print strVivodPredYarus + strCurrentList
    lngKolvoFormulPredYarus = CInt(Left(strVivodPredYarus, 1))
    If lngKolvoFormulPredYarus > 9 Then
        strVivodPredYarus = Replace(strVivodPredYarus, Left(strVivodPredYarus, 2) + "#", "", , 1)  'удалим первое число
    Else
        strVivodPredYarus = Replace(strVivodPredYarus, Left(strVivodPredYarus, 1) + "#", "", , 1)
    End If
    strObshSpisokForm = strVivodPredYarus 'Общий список формул ОбСпФо

    'Print #1, "Номер яруса формул", lngNYarus
    'Пока Nяр<3 и Длина(ВывЯрПред)<>0 делаем:
    Do While lngNYarus < 20 And lngKolvoFormulPredYarus <> 0
        Print #1, "' ############################################################################"
        Debug.Print "' ############################################################################"
        Debug.Print "Номер яруса формул", lngNYarus
        Print #1, "Номер яруса формул", lngNYarus 'Печать Nяр
        'В конце цикла вайл текущий ярус становится предыдущим и обновляются переменные.
        'В этом цикле выводятся формулы входящие в состав предыдущего яруса
        'и создается список переменных следующего яруса
        
        lngKolvoFormulTekYarus = 0
        lngIter = 1
        Do While lngIter < lngKolvoFormulPredYarus And StrComp(strVivodPredYarus, "") <> 0
            'Вывод описания формул текущей переменной в текущем ярусе
            'Print #1, "=====" + strVivodPredYarus
            strCurrentChar = Left(strVivodPredYarus, InStr(strVivodPredYarus, "#") - 1) 'Выделение названия текущей переменной
            strVivodPredYarus = Replace(strVivodPredYarus, strCurrentChar + "#", "", , 1) 'Удаляем название текущей переменной из списка формул
            strCurrentList = Left(strCurrentChar, InStr(strCurrentChar, "!") - 1) 'Выделение названия листа текущей переменной
            strCurrentChar = Replace(strCurrentChar, strCurrentList + "!", "", , 1) 'Удаляем названия листа текущей переменной
                       
            strHelp = strKolvoAndSpisokFormulCharov(strCurrentList, strCurrentChar, lngNYarus)
'            lngHelp = CInt(Left(strHelp, 1)) 'колво формул в текущей переменной
'            strHelp = Replace(strHelp, Left(strHelp, 1) + "#", "", , 1)  'удалим первое число
            
            If InStr(strHelp, "#") < 3 Then 'удалим первое число
                lngHelp = CInt(Left(strHelp, 1)) 'колво формул в текущей переменной
                strHelp = Replace(strHelp, Left(strHelp, 1) + "#", "", , 1)  'удалим первое число
            Else
                lngHelp = CInt(Left(strHelp, 2)) 'колво формул в текущей переменной
                strHelp = Replace(strHelp, Left(strHelp, 2) + "#", "", , 1)  'удалим первое число
            End If
                           
            If lngHelp > 0 Then
                lngKolvoFormulTekYarus = lngKolvoFormulTekYarus + lngHelp
                strVivodTekYarus = strVivodTekYarus + strHelp 'Вывод переменных в новый ярус
            End If
            'Создание списка переменных следующего яруса 'ВывЯр = ВывЯр + f(ВывЯрПред(i)) №1
            lngNForm = lngNForm + 1
        Loop 'Конец цикла
         
        strVivodTekYarus = strDeleter(strVivodTekYarus, strObshSpisokForm) '  ВывЯр = ВывЯр - ОбСпФо
        strVivodTekYarus = strCopyDeleter(strVivodTekYarus)
        strObshSpisokForm = strObshSpisokForm + strVivodTekYarus '  ОбСпФо = ОбСпФо + ВывЯр  №2
        strVivodTekYarus = CStr(lngKolvoFormulTekYarus) + "#" + strVivodTekYarus
        strVivodPredYarus = strVivodTekYarus 'ВывЯрПред = ВывЯр
        Print #1, "=====" + strVivodPredYarus
        If lngKolvoFormulTekYarus > 9 Then
            If lngKolvoFormulTekYarus > 99 Then
                lngKolvoFormulPredYarus = CInt(Left(strVivodPredYarus, 3)) 'обновляем кол-во формул предыдущего яруса
                strVivodPredYarus = Replace(strVivodPredYarus, Left(strVivodPredYarus, 3) + "#", "", , 1)   'удалим первое число
            Else
                lngKolvoFormulPredYarus = CInt(Left(strVivodPredYarus, 2)) 'обновляем кол-во формул предыдущего яруса
                strVivodPredYarus = Replace(strVivodPredYarus, Left(strVivodPredYarus, 2) + "#", "", , 1)   'удалим первое число
            End If
        Else
            lngKolvoFormulPredYarus = CInt(Left(strVivodPredYarus, 1)) 'обновляем кол-во формул предыдущего яруса
            strVivodPredYarus = Replace(strVivodPredYarus, Left(strVivodPredYarus, 1) + "#", "", , 1)
        End If

        strVivodTekYarus = "" 'ВывЯр = ""
        lngNYarus = lngNYarus + 1 'Nяр = Nяр + 1
    Loop 'Конец цикла

Close #1
 
End Sub

'функция добавляющая в накопитель переменные которых там нет
Function strCopyDeleter(strVivodYarus1)
    Dim strVarChar As String 'Помогалка для выделения формул
    Dim lngHelp As Long 'Помогалка для выделения формул
    Dim strHelp As String
    strHelp = strVivodYarus1

    Do While StrComp(strHelp, "") <> 0 And lngHelp < 300
        strVarChar = Left(strHelp, InStr(strHelp, "#")) 'Выделение названия текущей переменной
        strVivodYarus1 = Replace(strVivodYarus1, strVarChar, "$HEEEY$", , 1) 'Удаляем первое название текущей переменной из списка формул
        strVivodYarus1 = Replace(strVivodYarus1, strVarChar, "")
        strVivodYarus1 = Replace(strVivodYarus1, "$HEEEY$", strVarChar, , 1) 'Возвращаем первое название
        strHelp = Replace(strHelp, strVarChar, "") 'Удаляем название текущей переменной из накопителя
        lngHelp = lngHelp + 1
    Loop 'Конец цикла
    strCopyDeleter = strVivodYarus1
    
End Function

'функция удаляющая из яруса переменные накопителя (для удаления повторов из предыдущих ярусов)
Function strDeleter(strVivodYarus, strObshSpisokForm)
    Dim strVarChar As String 'Помогалка для выделения формул
    Dim lngHelp As Long 'Помогалка для выделения формул
    Dim strHelp As String
    strHelp = strObshSpisokForm

    Do While StrComp(strHelp, "") <> 0 And lngHelp < 300
        strVarChar = Left(strHelp, InStr(strHelp, "#")) 'Выделение названия текущей переменной
        strVivodYarus = Replace(strVivodYarus, strVarChar, "") 'Удаляем название текущей переменной из списка формул
        strHelp = Replace(strHelp, strVarChar, "") 'Удаляем название текущей переменной из накопителя
        lngHelp = lngHelp + 1
    Loop 'Конец цикла
    strDeleter = strVivodYarus
End Function


Function strAboutChar(strList, strChar, strType, strLeftNomerForm, lngHowMuchForm, Yarus)
    Dim strHelp As String 'Строковый Помогалка
    Dim strWhoIsIt As String 'Вывод названия переменной
    'strHelp = Worksheets(strList).Range(strChar).Row - 1 & ", " & Worksheets(strList).Range(strChar).Column 'Узнаем название переменной
   
    If Worksheets(strList).Cells(1, Worksheets(strList).Range(strChar).Column) = "" Then 'удалим первое число
        strWhoIsIt = Mid(strList, 1, 1) + Mid(strList, InStr(strList, " ") + 1, 1) + "Вспомогатель" + "_" + CStr(lngHowMuchForm) + "_" + CStr(Yarus) 'Узнаем название переменной
    Else
        strWhoIsIt = Mid(strList, 1, 1) + Mid(strList, InStr(strList, " ") + 1, 1) + Worksheets(strList).Cells(1, Worksheets(strList).Range(strChar).Column) + "_" + CStr(lngHowMuchForm) + CStr(Yarus) 'Узнаем название переменной
    End If
       
    'удаление пробелов /%$-.,()*
    strWhoIsIt = Replace(strWhoIsIt, " ", "")
    strWhoIsIt = Replace(strWhoIsIt, "/", "")
    strWhoIsIt = Replace(strWhoIsIt, "%", "")
    strWhoIsIt = Replace(strWhoIsIt, "$", "")
    strWhoIsIt = Replace(strWhoIsIt, "-", "")
    strWhoIsIt = Replace(strWhoIsIt, ".", "")
    strWhoIsIt = Replace(strWhoIsIt, ",", "")
    strWhoIsIt = Replace(strWhoIsIt, "(", "")
    strWhoIsIt = Replace(strWhoIsIt, ")", "")
    strWhoIsIt = Replace(strWhoIsIt, "*", "")
    strWhoIsIt = Replace(strWhoIsIt, "Chr(34)", "")
    strWhoIsIt = Replace(strWhoIsIt, vbNewLine, "")
    
    If strWhoIsIt = "" Then
        strWhoIsIt = "Ячейка без названия"
    End If

    If strType = "Формула" Then 'Если формула, то добавляем в конце номер, иначе нет
        strAboutChar = "Dim " + strWhoIsIt + " As Double" + " ' " + strType + " " + CStr(lngHowMuchForm) + " " + strList + " " + strChar + vbNewLine
        strAboutChar = strAboutChar + strWhoIsIt + " = " + Replace(CStr(Worksheets(strList).Cells(2, Worksheets(strList).Range(strChar).Column)), ",", ".")
    ElseIf strType = "Просто объява" Then
        strAboutChar = "Dim " + strWhoIsIt + " As Double" + " ' " + " " + " " + strList + " " + strChar
    ElseIf strType = "Текущая Формула" Then
        strAboutChar = " " + strList + " " + strChar + " " + " = " + Replace(CStr(Worksheets(strList).Cells(2, Worksheets(strList).Range(strChar).Column)), ",", ".")
    ElseIf strType = "Просто название" Then
        strAboutChar = strWhoIsIt
    Else
        strAboutChar = "Dim " + strWhoIsIt
        strAboutChar = strAboutChar + " As Double" + " ' " + strType + " " + strList + " " + strChar + vbNewLine
        strAboutChar = strAboutChar + strWhoIsIt + " = " + Replace(CStr(Worksheets(strList).Cells(2, Worksheets(strList).Range(strChar).Column)), ",", ".")
    End If
End Function


Function strKolvoAndSpisokFormulCharov(strCurrentList, strCurrentChar, lngNYarus) 'поставить курент формула входным или внутрь
    Dim strHelp As String 'Строковый Помогалка
    Dim strOutputPrintCurrentFormula As String 'Строковый Помогалка
    Dim lngHelp As String 'Целочисленный Помогалка
    Dim lngHelp2 As String 'Целочисленный Помогалка
    Dim lngIterChar As Long 'Целочисленный итератор по таблице
    Dim strIter As Variant 'Строковый итератор
    Dim strMainFormula As String 'Формула
    Dim strIterChar As String 'Переменная для подстановки ячейки
    Dim arrSearchField(1 To 2, 1 To 1) As Variant 'Область поиска
    Dim lngHowMuchForm As Long 'Счетчик формул
    Dim lngWhereIterCharInFormula As Long 'Где в формуле ячейка
    Dim strCurrentFormula As String 'Текущая формула на обработке
    Dim lngWhereListIterCharInFormula As Long 'Константа для поиска в листа в формуле
    Dim strListIterChar As String 'Текущая формула на обработке
    'Dim strKolvoAndSpisokFormulCharov As String 'Выходное
    Dim strType As String 'Тип переменной формула или число
    'Dim strLeftNomerForm As String 'Строковый Помогалка
    
    Dim strLeftNomerForm As String
    
    Print #1, "' ############################################################################"
    Debug.Print "' ############################################################################"
    
    
    
    strCurrentFormula = ThisWorkbook.Worksheets(strCurrentList).Range(strCurrentChar).Formula
    strCurrentFormula = Replace(strCurrentFormula, "$", "")
    strCurrentFormula = Replace(strCurrentFormula, "IF", "IIf")
    strHelp = strAboutChar(strCurrentList, strCurrentChar, "Текущая Формула", strLeftNomerForm, 0, lngNYarus)
    strOutputPrintCurrentFormula = strCurrentFormula
    
    Debug.Print strAboutChar(strCurrentList, strCurrentChar, "Просто объява", strLeftNomerForm, 0, lngNYarus)
    Print #1, strAboutChar(strCurrentList, strCurrentChar, "Просто объява", strLeftNomerForm, 0, lngNYarus)

      
    strKolvoAndSpisokFormulCharov = ""
    lngHowMuchForm = 0
    For lngIterChar = 300 To 1 Step -1 'проходка по ячейкам AW2..A2
        lngWhereIterCharInFormula = 0
        strIterChar = Worksheets(strCurrentList).Cells(2, lngIterChar).Address 'перевод численного адреса на ячейку в буквенный
        strIterChar = Replace(strIterChar, "$", "") 'убираем доллары из названия адреса
        
        lngWhereIterCharInFormula = InStr(strCurrentFormula, strIterChar) 'где буквенный адрес в формуле
        If lngWhereIterCharInFormula > 0 Then 'если буквенный адрес есть в формуле
        
            lngIterChar = lngIterChar + 1
            If Mid(strCurrentFormula, lngWhereIterCharInFormula - 1, 1) = "!" Then 'если есть ссылка на лист, то обрабатываем ее
                
                lngWhereListIterCharInFormula = InStrRev(strCurrentFormula, "'", lngWhereIterCharInFormula - 3)
                strListIterChar = Mid(strCurrentFormula, lngWhereListIterCharInFormula + 1, lngWhereIterCharInFormula - lngWhereListIterCharInFormula - 3)
                If InStr(Worksheets(strListIterChar).Range(strIterChar).Formula, "=") = 1 Then 'проверка является ли формулой
                    strType = "Формула"
                    lngHowMuchForm = lngHowMuchForm + 1
                    Print #1, strAboutChar(strListIterChar, strIterChar, strType, strLeftNomerForm, lngHowMuchForm, lngNYarus) 'Вывод информации о переменной
                    Debug.Print strAboutChar(strListIterChar, strIterChar, strType, strLeftNomerForm, lngHowMuchForm, lngNYarus)
                    strOutputPrintCurrentFormula = Replace(strOutputPrintCurrentFormula, "'" + strListIterChar + "'" + "!" + strIterChar, strAboutChar(strListIterChar, strIterChar, "Просто название", strLeftNomerForm, lngHowMuchForm, lngNYarus), , 1)
                    strKolvoAndSpisokFormulCharov = strKolvoAndSpisokFormulCharov + strListIterChar + "!" + strIterChar + "#"
                    strCurrentFormula = Replace(strCurrentFormula, "'" + strListIterChar + "'" + "!" + strIterChar, "", , 1)
                Else
                    lngHowMuchForm = lngHowMuchForm + 1 'увеличиваем счетчик формул
                    strType = "Константа"
                    Print #1, strAboutChar(strListIterChar, strIterChar, strType, 0, lngHowMuchForm, lngNYarus) 'Вывод информации о переменной
                    Debug.Print strAboutChar(strListIterChar, strIterChar, strType, 0, lngHowMuchForm, lngNYarus)
                    strOutputPrintCurrentFormula = Replace(strOutputPrintCurrentFormula, "'" + strListIterChar + "'" + "!" + strIterChar, strAboutChar(strListIterChar, strIterChar, "Просто название", strLeftNomerForm, lngHowMuchForm, lngNYarus), , 1)
                    strCurrentFormula = Replace(strCurrentFormula, "'" + strListIterChar + "'" + "!" + strIterChar, "", , 1)
                   
                End If

            Else 'если нет ссылка на лист, то не обрабатываем ее
                strListIterChar = strCurrentList
                If InStr(Worksheets(strCurrentList).Range(strIterChar).Formula, "=") = 1 Then 'проверка является ли формулой
                    lngHowMuchForm = lngHowMuchForm + 1 'увеличиваем счетчик формул
                    strType = "Формула"
                    Print #1, strAboutChar(strListIterChar, strIterChar, strType, strLeftNomerForm, lngHowMuchForm, lngNYarus) 'Вывод информации о переменной
                    Debug.Print strAboutChar(strListIterChar, strIterChar, strType, strLeftNomerForm, lngHowMuchForm, lngNYarus)
                    strKolvoAndSpisokFormulCharov = strKolvoAndSpisokFormulCharov + strListIterChar + "!" + strIterChar + "#"
                    strOutputPrintCurrentFormula = Replace(strOutputPrintCurrentFormula, strIterChar, strAboutChar(strListIterChar, strIterChar, "Просто название", strLeftNomerForm, lngHowMuchForm, lngNYarus), , 1)
                    strCurrentFormula = Replace(strCurrentFormula, strIterChar, "", , 1)
                Else
                    lngHowMuchForm = lngHowMuchForm + 1 'увеличиваем счетчик формул
                    strType = "Константа"
                    Print #1, strAboutChar(strListIterChar, strIterChar, strType, 0, lngHowMuchForm, lngNYarus) 'Вывод информации о переменной
                    Debug.Print strAboutChar(strListIterChar, strIterChar, strType, 0, lngHowMuchForm, lngNYarus)
                    strOutputPrintCurrentFormula = Replace(strOutputPrintCurrentFormula, strIterChar, strAboutChar(strListIterChar, strIterChar, "Просто название", strLeftNomerForm, lngHowMuchForm, lngNYarus), , 1)
                    'strOutputPrintCurrentFormula = Replace(strOutputPrintCurrentFormula, strIterChar, strAboutChar(strListIterChar, strIterChar, "Просто название", strLeftNomerForm, lngHowMuchForm))
                    strCurrentFormula = Replace(strCurrentFormula, strIterChar, "", , 1)
                End If
            End If
            
        End If
    Next lngIterChar
    Print #1, strAboutChar(strCurrentList, strCurrentChar, "Просто название", strLeftNomerForm, 0, lngNYarus) + strOutputPrintCurrentFormula
    Print #1, " ' " + strHelp
    Print #1, "Debug.Print " + strAboutChar(strCurrentList, strCurrentChar, "Просто название", strLeftNomerForm, 0, lngNYarus)
    Debug.Print strAboutChar(strCurrentList, strCurrentChar, "Просто название", strLeftNomerForm, 0, lngNYarus) + strOutputPrintCurrentFormula
    Debug.Print " ' " + strHelp
    Debug.Print "Debug.Print " + strAboutChar(strCurrentList, strCurrentChar, "Просто название", strLeftNomerForm, 0, lngNYarus)
    strKolvoAndSpisokFormulCharov = CStr(lngHowMuchForm) + "#" + strKolvoAndSpisokFormulCharov 'запись в переменную формул
End Function




