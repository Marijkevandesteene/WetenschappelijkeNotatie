Attribute VB_Name = "Wetenschappelijk"
Option Explicit
Public Const cStrFuncties = "sin;cos;tan;cot;sec;csc"
Public Const cNbrSets = "R;N;Q;N;Z;I"
Public Const cStrSymbolen = "infty;geq;leq"
Public Const cKarakters = "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z"

Public Function IsInArray(ByVal vToFind As Variant, vArr As Variant) As Boolean

    Dim i As Long
    Dim bReturn As Boolean
    Dim vLine As Variant
    
    bReturn = False
    
    For i = 0 To UBound(vArr, 1)
        If vToFind = vArr(i) Then
            bReturn = True
        End If
        If bReturn Then Exit For 'stop looking if one found
    Next i

    IsInArray = bReturn

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Symbool
Function Symbool(sCode As String, nbrWord As Integer) As Boolean
'
' Functie voor het invoeren van symbool met LaTeX-code sCode
'
    Dim objrange As Range
    Dim objEq As OMath
     
    Set objrange = Selection.Range
    objrange.Text = sCode
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
    
    Symbool = True

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' NaSelectie
Function NaSelectie(sCode As String, nbrWord As Integer) As Boolean
'
' Functie voor het invoeren van codeneNaSelectie met LaTeX-code sCode
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formuleTekst As String
                  
    Set objrange = Selection.Range
    If Selection.Type = wdSelectionIP Then
        objrange.Text = objrange.Text & sCode
        Set objrange = Selection.OMaths.Add(objrange)
    Else
        Set objrange = Selection.OMaths.Add(objrange)
        Set objEq = objrange.OMaths(1)
         objEq.Linearize
        formuleTekst = Selection.Text
        formuleTekst = formuleTekst & sCode & " "
        formuleTekst = Replace(formuleTekst, "\\", "\\ ")
        'Selection.Text = formuleTekst
        objrange.Text = formuleTekst
        Set objrange = Selection.OMaths.Add(objrange)
    End If

 '   objRange.Text = objRange.Text & sCode
 
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
    
    NaSelectie = True

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' VoorSelectie
Function voorSelectie(sCode As String, nbrWord As Integer) As Boolean
'
' Functie voor het invoeren van codeneNaSelectie met LaTeX-code sCode
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formuleTekst As String
                  
    Set objrange = Selection.Range
    If Selection.Type = wdSelectionIP Then
        objrange.Text = sCode & objrange.Text
        Set objrange = Selection.OMaths.Add(objrange)
    Else
        Set objrange = Selection.OMaths.Add(objrange)
        Set objEq = objrange.OMaths(1)
         objEq.Linearize
        formuleTekst = Selection.Text
        formuleTekst = sCode & formuleTekst & " "
        formuleTekst = Replace(formuleTekst, "\\", "\\ ")
        Selection.Text = formuleTekst
    End If

 '   objRange.Text = objRange.Text & sCode
 
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
     
    voorSelectie = True

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' VanSelectie
Function VanSelectie(sCode As String, nbrWord As Integer) As Boolean
'
' Functie voor het invoeren van functie van Selectie met LaTeX-code sCode
'
    Dim objrange As Range
    Dim objEq As OMath
     
    Set objrange = Selection.Range
    objrange.Text = "\" & sCode & "{" & objrange.Text & "}"
 
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
    
    VanSelectie = True

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Algemeen
Sub Breuk()
'
' Breuk Macro
'
    Dim gelukt As Boolean
    Dim objrange As Range
    Dim objEq As OMath
    Dim lengte As Integer
    Dim delen
    
    Set objrange = Selection.Range
    If Selection.Type = wdSelectionIP Then
        objrange.Text = "\frac{}{}"
    Else
        If InStr(objrange.Text, "/") > 0 Then
            delen = Split(objrange.Text, "/")
            lengte = UBound(delen)
            If lengte = 1 Then
                objrange.Text = "\frac{" & delen(0) & "}{" & delen(1) & "}"
            ElseIf lengte = 0 Then
               objrange.Text = "\frac{" & delen(0) & "}{}"
            Else
                objrange.Text = "\frac{" & objrange.Text & "}{}"
            End If
        End If
     End If
    
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
   

End Sub
Sub Vierkantswortel()
'
' Vierkantswortel Macro
'
    Dim gelukt As Boolean
'    gelukt = Symbool("\sqrt{}", 1)
    gelukt = VanSelectie("sqrt", 1)
    
End Sub
Sub Wortel()
'
' Wortel Macro
'
    Dim gelukt As Boolean
    'gelukt = Symbool("\sqrt[n]{}", 3)
    gelukt = VanSelectie("sqrt[n]", 1)
   
End Sub
Sub VoegFunctieToe()
'
' VoegFunctieToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sFunctie As String
    sFunctie = InputBox("Goniometrische functies: sin, cos, tan, ... " & Chr(13) & "Andere functies: adj, det" & Chr(13) & "Voor vector: vec" & Chr(13) & "Andere: overline, underline", "Welke functie wil je invoeren", "sin")
     
    Set objrange = Selection.Range
    objrange.Text = "\" & sFunctie & "{" & objrange.Text & "}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub
Sub VoegSymboolToe()
'
' VoegSymboolToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim bericht, sSymbool As String
    Dim Symbolen, NbrSets
    NbrSets = Split(cNbrSets, ";")
    Symbolen = Split(cStrSymbolen, ";")
    bericht = "Voor oneindig: infty" & Chr(13) & "Voor Griekse letters: alpha, beta, ... " & Chr(13) & "Voor Griekse hoofdletters: Alpha, Beta, ..., Delta " & Chr(13) & "Voor >=: geq, voor <=: leq" & Chr(13) & "Voor getallenverzamelingen: N, Z, Q, R, I"
    
    sSymbool = InputBox(bericht, "Welk symbool?", "infty")
     
    Set objrange = Selection.Range
    objrange.Text = "\" & sSymbool
    If sSymbool = "oneindig" Then
        objrange.Text = "\infty"
    End If
    If sSymbool = ">=" Then
        objrange.Text = "\geq"
    End If
    If sSymbool = "<=" Then
        objrange.Text = "\leq"
    End If
    If IsInArray(sSymbool, NbrSets) Then
         objrange.Text = "\mathbb{\double" & sSymbool & "}"
    End If
    If IsInArray(sSymbool, Symbolen) Then
         objrange.Text = "\" & sSymbool
    End If
   
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

Function BepaalGrenzen4Integraal(sGrens)
    Dim grens As String
    
    If Len(sGrens) = 1 Then
        grens = sGrens
    ElseIf Mid(sGrens, 1, 1) = "-" Then
        'grens = "-"
        If IsNumeric(Mid(sGrens, 2, 1)) Then
            grens = sGrens
        Else
            grens = "-" & "\" & Mid(sGrens, 2, Len(sGrens) - 1)
        End If
    ElseIf Mid(sGrens, 1, 1) = "+" Then
        'grens = "+"
        If IsNumeric(Mid(sGrens, 1, 1)) Then
            grens = sGrens
        ElseIf Len(sGrens) = 2 Then
            grens = "+" & Mid(sGrens, 2, Len(sGrens) - 1)
        Else
            grens = "+" & "\" & Mid(sGrens, 2, Len(sGrens) - 1)
        End If
    Else
        If IsNumeric(Mid(sGrens, 1, 1)) Then
            grens = sGrens
        Else
            grens = "\" & sGrens
        End If
        
    End If
    BepaalGrenzen4Integraal = grens
 
End Function
Sub VoegBepaaldeIntegralenToe()
'
' VoegSymboolToe Macro
'
    Dim objrange, objOut As Range
    Dim objEq As OMath
    Dim graad As Integer
    Dim sVan, sTot As String
    Dim van, tot As String
    
    graad = InputBox("Welke graad heeft de integraal?", "Voeg een bepaalde integraal toe", "1")
    sVan = InputBox("Geef de ondergrens voor de bepaalde integraal. Voor een symbool gebruik de code, bvb -infty voor min ondeindig", "Ondergrens van integraal?", "0")
    sTot = InputBox("Geef de bovengrens voor de bepaalde integraal. Voor een symbool gebruik de code, bvb +z voor +z", "Bovengrens van integraal?", "infty")
    
    Set objrange = Selection.Range
    van = BepaalGrenzen4Integraal(sVan)
    tot = BepaalGrenzen4Integraal(sTot)
    
    If graad = 1 Then
        objrange.Text = "\int^{" & tot & "}_{" & van & "}{" & objrange.Text & "}"
    ElseIf graad = 2 Then
        objrange.Text = "\iint^{" & tot & "}_{" & van & "}{" & objrange.Text & "}"
    ElseIf graad = 3 Then
        objrange.Text = "\iiint^{" & tot & "}_{" & van & "}{" & objrange.Text & "}"
    End If

    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub
Sub Limiet()
'
' limiet Macro
'
    Dim objrange, objOut As Range
    Dim objEq As OMath
    Dim graad As Integer
    Dim sVan, sTot As String
    Dim van, tot As String
    
    sVan = InputBox("Geef de limiet voor variabele ***. Voor een symbool gebruik de code, bvb alpha", "Limiet voor ***?", "x")
    sTot = InputBox("Geef de limiet gaande naar de bovengrens ***. Voor een symbool gebruik de code, bvb infty voor +oneindig", "Limiet tot *?", "infty")
    
    Set objrange = Selection.Range
    van = BepaalGrenzen4Integraal(sVan)
    tot = BepaalGrenzen4Integraal(sTot)
    
    objrange.Text = "\lim_{" & van & "\rightarrow" & tot & "}{" & objrange.Text & "}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

Sub Afgeleide()
'
' limiet Macro
'
    Dim objrange, objOut As Range
    Dim objEq As OMath
         
    Set objrange = Selection.Range
    
    objrange.Text = "\frac{d(" & objrange.Text & ")}{dx}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
End Sub


Sub VoegOnbepaaldeIntegralenToe()
'
' VoegSymboolToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim graad As Integer
    
    graad = InputBox("Welke graad heeft de integraal?", "Voeg een onbepaalde integraal toe", "1")
    
    Set objrange = Selection.Range
    
    If graad = 1 Then
        objrange.Text = "\int{" & objrange.Text & "}"
    ElseIf graad = 2 Then
        objrange.Text = "\iint{" & objrange.Text & "}"
    ElseIf graad = 3 Then
        objrange.Text = "\iiint{" & objrange.Text & "}"
    End If

    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

Sub VoegSomToe()
'
' VoegSymboolToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim sVan, sTot As String
    
    'graad = InputBox("Welke graad heeft de integraal?", "Voeg een bepaalde integraal toe", "1")
    sVan = InputBox("Geef het start interval voor de som van een functie.", "Eerste deelinterval", "i = 1")
    sTot = InputBox("Geef het aantal deelintervallen voor de som van de functie.", "Aantal deelintervallen?", "n")
    
    Set objrange = Selection.Range
    
    objrange.Text = "\sum^{" & sTot & "}_{" & sVan & "}{" & objrange.Text & "}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

''Macht en Indices
Sub Macht()
'
' Macht Macro
'
    Dim gelukt As Boolean
    gelukt = NaSelectie("^{}", 1)

End Sub
Sub Index()
'
' Index Macro
'
    Dim gelukt As Boolean
    gelukt = NaSelectie("_{}", 1)
       
End Sub
Sub MachtEnIndex()
'
' MachtEnIndex Macro
'
    Dim gelukt As Boolean
    gelukt = NaSelectie("^{}_{}", 1)

End Sub

Sub boven()
'
' Boven Macro
'
    Dim gelukt As Boolean
    gelukt = NaSelectie("\above{}", 1)

End Sub
Sub onder()
'
' Onder Macro
'
    Dim gelukt As Boolean
    gelukt = NaSelectie("\below{}", 1)

End Sub

Sub Atoomgetal()
'
' Atoomgetal Macro Z
'
    Dim gelukt As Boolean
    Dim objrange As Range
    Dim objEq As OMath
    Dim lengte As Integer
    Dim delen
    
    Set objrange = Selection.Range
    'objRange.Text = Selection.Text
    Set objrange = Selection.OMaths.Add(objrange)
    objrange.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
    wdOMathFunctionScrPre
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp
    
End Sub

Sub overline()
'
' VertikaleStreep Macro
'
    Dim gelukt As Boolean
    gelukt = VanSelectie("overline", 1)

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Haakjes
Function Haakjes(links As String, rechts As String) As Boolean
'
' Functie voor het invoeren van haakjes links en rechts van een selectie of zonder inhoud
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formuleTekst, tName As String
         
    'ga 1karakter (newline) terug om geen paragraf mee te hebben
    Set objrange = Selection.Range
    
      
    tName = typeName(Selection)
    'check if selection is empty
    If Selection.Type = wdSelectionIP Then
        objrange.Text = links & rechts
        Set objrange = Selection.OMaths.Add(objrange)
    Else
        Set objrange = Selection.OMaths.Add(objrange)
        Set objEq = objrange.OMaths(1)
        objEq.Linearize
        formuleTekst = Selection.Text
        formuleTekst = Replace(formuleTekst, "\\", "\\ ")
       Selection.Text = links & formuleTekst & rechts
    End If
    Set objEq = objrange.OMaths(1)
    
    objEq.BuildUp
    Haakjes = True

End Function
Sub VierkanteHaakjes()
'
' VierkanteHaakjes Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left[", "\right]")

End Sub
Sub RondeHaakjes()
'
' RondeHaakjes Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left(", "\right)")

End Sub
Sub Accolades()
'
' Accolades Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left{", "\right}")

End Sub

Sub Accolade()
'
' Accolade Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left{", "\right.")

End Sub

Sub VertikaleStrepen()
'
' VertikaleStrepen Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left|", "\right|")

End Sub

Sub VertikaleStreep()
'
' VertikaleStreep Macro
'
    Dim gelukt As Boolean
    gelukt = Haakjes("\left.", "\right|")
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Pijlen
Sub naar()
'
' naar Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\longrightarrow", 1)

End Sub
Sub van()
'
' van Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\longleftarrow", 1)

End Sub
Sub vanEnNaar()
'
' vanEnNaar Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\longleftrightarrow", 1)

End Sub
Sub naar2()
'
' naar Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\Longrightarrow", 1)

End Sub
Sub van2()
'
' van Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\Longleftarrow", 1)

End Sub
Sub vanEnNaar2()
'
' vanEnNaar2 Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\Longleftrightarrow", 1)

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' VaakGebruikt
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Functies voor maken van Matrix, stelsel
'

Function MaakMxN_Rooster(ByVal m As Integer, N As Integer) As String
    Dim formule As String
    formule = ""
    Dim rij, kolom As Integer
    
    For rij = 1 To m
        For kolom = 1 To N - 1
            formule = formule + "&"
        Next kolom
        formule = formule + "\\ "
    Next rij
    MaakMxN_Rooster = formule
    
End Function
Function MaakMx_Lijst(ByVal m As Integer) As String
    Dim formule As String
    Dim element As Integer
    
    formule = ""
    For element = 1 To m
        formule = formule + "\\ "
    Next element
    MaakMx_Lijst = formule
    
End Function

Sub Logaritme()
'
' Logaritme Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("^a\log{}", 5)

End Sub
Sub ln()
'
' ln Macro
'
    Dim gelukt As Boolean
    gelukt = Symbool("\ln{}", 3)

End Sub
Sub VoegMxNroosterToe()
'
' Voeg1x2matrixToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sRijen, sKolommen As String
    sRijen = InputBox("Hoeveel rijen?", "#rijen", "2")
    sKolommen = InputBox("Hoeveel kolommen?", "#kolommen", "2")
    Dim m, N As Integer
    N = Int(sKolommen)
    m = Int(sRijen)
     
    Set objrange = Selection.Range
    formule = MaakMxN_Rooster(m, N)
    objrange.Text = "\begin{matrix}" & formule & "\end{matrix}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub
Sub VoegMxNmatrixToe()
'
' Voeg1x2matrixToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sRijen, sKolommen As String
    sRijen = InputBox("Hoeveel rijen?", "#rijen", "2")
    sKolommen = InputBox("Hoeveel kolommen?", "#kolommen", "2")
    Dim m, N As Integer
    N = Int(sKolommen)
    m = Int(sRijen)
     
    Set objrange = Selection.Range
    formule = MaakMxN_Rooster(m, N)
    objrange.Text = "\left[\begin{matrix}" & formule & "\end{matrix}\right]"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub
Sub VoegMxNdeterminantToe()
'
' VoegMxNdeterminantToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sRijen, sKolommen As String
    sRijen = InputBox("Hoeveel rijen?", "#rijen", "2")
    sKolommen = InputBox("Hoeveel kolommen?", "#kolommen", "2")
    Dim m, N As Integer
    N = Int(sKolommen)
    m = Int(sRijen)
     
    Set objrange = Selection.Range
    formule = MaakMxN_Rooster(m, N)
    objrange.Text = "\left|\begin{matrix}" & formule & "\end{matrix}\right|"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub
Sub VoegUitgebreideMatrixToe()
'
' Voeg1x2matrixToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, matrix, kolommatrix, sRijen, sKolommen As String
    sRijen = InputBox("Hoeveel rijen zijn er?", "#kolommen", "2")
    sKolommen = InputBox("Hoeveel kolommen zijn er in de uitgebreide matrix?", "#rijen", "3")
    Dim m, N As Integer
    N = Int(sKolommen)
    m = Int(sRijen)
     
    Set objrange = Selection.Range
    matrix = "\begin{matrix}" & MaakMxN_Rooster(m, N) & "\end{matrix}"
    kolommatrix = "\begin{matrix}" & MaakMx_Lijst(m) & "\end{matrix}"
    formule = matrix & "\left|" & kolommatrix & "\right."
    formule = "\left[" & formule & "\right]"
    objrange.Text = formule
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

Sub VoegMxVergelijkingToe()
'
' VoegMxVergelijkingToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sVergelijkingen As String
    sVergelijkingen = InputBox("Hoeveel vergelijkingen?", "#Vergelijkingen", "2")
    Dim m As Integer
    m = Int(sVergelijkingen)
     
    Set objrange = Selection.Range
    formule = MaakMx_Lijst(m)
    objrange.Text = "\left{\begin{matrix}" & formule & "\end{matrix}\right."
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

Sub VoegMxVectorToe()
'
' VoegMxVergelijkingToe Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim formule, sElementen As String
    sElementen = InputBox("Hoeveel elementen?", "#elementen", "2")
    Dim m As Integer
    m = Int(sElementen)
     
    Set objrange = Selection.Range
    formule = MaakMx_Lijst(m)
    objrange.Text = "\begin{matrix}" & formule & "\end{matrix}"
    
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Tools

Sub NaarWetenschappelijk()
'
' NaarWetenschappelijk Macro
' '//' vervangen door '// ' om niet als extra element vertaald te worden
'
    Dim objrange As Range
    Dim objEq, objItem As OMath
       
    If Selection.Type = wdSelectionIP Then
        Set objrange = Selection.Range
        objrange.OMaths(1).Linearize
        
        objrange.Text = Replace(objrange.Text, "\ ", "")
        objrange.Text = Replace(objrange.Text, "\\", "\\ ")
        Set objrange = Selection.OMaths.Add(objrange)
        
        Set objEq = objrange.OMaths(1)
        objEq.BuildUp
    End If

End Sub

Sub Code2Wetenschappelijk()
'
' Code2Wetenschappelijk Macro
'
'
    Dim objrange, objItem As Range
    Dim objEq As OMath
    Dim sInputText As String
    
    Set objrange = Selection.Range
    If Selection.Type = wdSelectionIP Then
        objrange.Text = InputBox("Welke Code(LaTeX) wil je vertalen? _ voor 1 karakter in subscript, ^ voor 1 karakter in superscript, _{} of ^{} voor meerdere", "Welk LaTeX code?", "6 H_2O + 6CO_2 \to C_6H_{12}O_6 + 6 O_2")
        Set objrange = Selection.OMaths.Add(objrange)
    Else
         objrange.Text = Selection.Text
         Set objrange = Selection.OMaths.Add(objrange)
    End If

   Set objEq = objrange.OMaths(1)
        objEq.BuildUp
   
   

End Sub

Sub NaarCode()
'
' NaarCode Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim i As Integer
    
     
    Set objrange = Selection.Range
    Set objrange = Selection.OMaths.Add(objrange)
    For i = 1 To Selection.OMaths.Count
        Set objEq = objrange.OMaths(i)
        objEq.Linearize
    Next i

End Sub

Sub GewoneTekst()
'
' GewoneTekst Macro
'
    Dim objrange As Range
    Dim objEq As OMath
    Dim functies
    functies = Split(cStrFuncties, ";")

    Set objrange = Selection.Range
    objrange.Text = "\mathrm{" & objrange.Text & "}"
    Set objrange = Selection.OMaths.Add(objrange)
    Set objEq = objrange.OMaths(1)
    objEq.BuildUp

End Sub


