Attribute VB_Name = "NewMacros"

Sub abner_carta()
'
    Application.ScreenUpdating = False
    Dim doc As Document
    Set doc = Documents.Add
    Dim espacio As String
    Dim fuente As String
    Dim fuente_t As String
    Dim fuente_s As String
    Dim fuente_s_t As String
    Dim fuente_asunto_tam As String
    Dim fuente_asunto_nom As String
    fuente_asunto_nom = "Courier Final Draft"
    fuente_asunto_tam = 12
    fuente = "Courier Final Draft"
    fuente_t = 12
    'numCarxPulgada = 20 '--------- importante cuando la fuente es de paso variable 11.75 Lucida Bright (12), 13.5 Garamond (12)
    numEspaciosxPulgada = 12
    'k = fuente_asunto_tam * numCarxPulgada
    fuente_s = "Courier Final Draft"
    fuente_s_t = 12
    espacio = wdLineSpace1pt5
    
    With doc.range
        .Font.Name = fuente
        .Font.Size = fuente_t
        .ParagraphFormat.Hyphenation = True
        .PageSetup.TopMargin = InchesToPoints(1)
        .PageSetup.BottomMargin = InchesToPoints(1)
        .PageSetup.LeftMargin = InchesToPoints(1)
        .PageSetup.RightMargin = InchesToPoints(1)
        .PageSetup.FooterDistance = InchesToPoints(0.5)
        .PageSetup.HeaderDistance = InchesToPoints(0.5)
        .Document.AutoHyphenation = True
        .Document.HyphenateCaps = True
        .Document.HyphenationZone = InchesToPoints(0.3)
        .PageSetup.DifferentFirstPageHeaderFooter = True
    End With
    Dim s As Selection
    'Dim s As range
    Set s = doc.ActiveWindow.Selection
    Dim p As range
    '
    'Set p = s.Paragraphs(1).range
    'Set p = doc.range(doc.Paragraphs(1).range.Start, _
    doc.Paragraphs(doc.Paragraphs.Count).range.End)
    'Set p = doc.Paragraphs(1).range
    Set p = doc.range
    Dim Nombre As String
    Dim el As String
    Dim asunto As String
    Dim vocativo As String
    Dim firma As String
    Dim Continuar As Boolean
    Nombre = InputBox("Por favor, escriba el nombre para quien va dirigida la carta," _
            & " no escriba signos de puntuación al final. Como el nombre de la persona " _
            & "a quien envía la carta está asociado al nombre del archivo por crearse, no escriba " _
            & "los siguientes símbolos : \ * < > / ? """, "ANTES DE CONTINUAR...", "Destinatario")
    el = InputBox("Por favor, escriba ""el"" o ""la"" si antepuso al nombre algún " _
            & "título o tratamiento de cortesía como Lic., Sr. o Sra., si no es así, " _
            & "no escriba nada y dé clic en cancelar o aceptar, no escriba signos de " _
            & "puntuación o espacios al final del texto que escriba", "ANTES DE CONTINUAR...", "")
    asunto = InputBox("Por favor, escriba el asunto sobre el que versa la" _
            & " carta", "ANTES DE CONTINUAR...", "Saludos cordiales, etc.")
    lugar = InputBox("Por favor, escriba el lugar donde redacta su carta" _
            & " No coloque signos de puntuación.", "ANTES DE CONTINUAR...", "Ciudad de México")
    vocativo = InputBox("Por favor, escriba el vocativo para quien va dirigida la carta, " _
            & "no escriba signos de puntuación ni espacios (ver modelo por default", "ANTES" _
            & "DE CONTINUAR...", "Estimado")
    firma = InputBox("Por favor, escriba quien firma la presente la carta:" _
            & "", "ANTES DE CONTINUAR... ", "Nombre de la persona que firma este documento")
    'resp = IIf(MsgBox("¿Quiere continuar con la carta?" & _
                (InputBox("Por favor, escriba quien firma la presente la carta:" _
            & "", "ANTES DE CONTINUAR... ", "Nombre de la persona que firma este documento")), vbQuestion + vbOKCancel, "Atención") = vbYes, True, False)
            
    'resp = IIf(MsgBox("¿Quiere continuar con la carta para " & Nombre & "?" & _
                (InputBox("Por favor, escriba quien firma la presente la carta:" _
            & "", "ANTES DE CONTINUAR... ", "Nombre de la persona que firma este documento")), vbQuestion + vbYesNo, "Atención") = vbYes, True, False)
            
    Continuar = IIf(MsgBox("¿Desea continuar escribiendo la carta para el " & Nombre & "?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes, True, False)
    If Continuar = True Then
       GoTo 8
    End If
    If Continuar = False Then
       GoTo 9
    End If
8   mensajeinicial = "AVISO: Inicie desde esta zona." _
            & " Este párrafo sombreado tiene el espacio que usted haya elegido sin sangría de primera línea," _
            & " el siguiente párrafo conserva el mismo espaciado y añade una sangría de 0.5 pulgadas en" _
            & " la sangría de primera línea. En los párrafos de firma, se regresa a los párrafos con espacio simple." _
            & " Pulse alguna tecla de flecha para conservar este aviso o cualquier otra para que desaparezca."
    'valor = Format(Format(Now, "dddd"), "=") & " " & Format(Now, "d") & " de " & Format(Format(Now, "mmmm"), "<") & " de " & Format(Now, "yyyy") & " ¤ " & Format(Now(), "h·mm·ss") & _
            " • " & "Carta para «" & Nombre & "»" & ".docx"
    valor = Format(Format(Now, "dddd"), "=") & " " & Format(Now, "d") & " de " & Format(Format(Now, "mmmm"), "<") & " de " & Format(Now, "yyyy") & " @" & Format(Now(), "h-mm-ss") & "h" & _
            "  " & "Carta para " & Nombre & ".docx"
        
    If Not Dialogs(wdDialogFileSaveAs).Display Then Exit Sub
        Dialogs(wdDialogFileSaveAs).Application.ActiveDocument.SaveAs2 FileName:=valor, FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
    If Dialogs(wdDialogFileSaveAs) Then
    End If
    
    With doc.Sections(1).range
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .ParagraphFormat.FirstLineIndent = InchesToPoints(0)
            .Font.Name = fuente
            .Font.Size = fuente_t
            .PageSetup.HeaderDistance = InchesToPoints(0.5)
            .PageSetup.FooterDistance = InchesToPoints(0.5)
    End With
    With doc.ActiveWindow.View
        .SeekView = wdSeekFirstPageHeader
            b1 = " y uno segundos."
            r1 = " y un segundos."
            b2 = " a. m. "
            r2 = " mañana "
            b3 = " p. m. "
            r3 = " tarde "
            b4 = " p.m. "
            r4 = " tarde "
            b5 = " a.m. "
            r5 = " mañana "
            b6 = "Doce de la tarde"
            r6 = "Doce del día"
            b7 = "Siete de la tarde"
            r7 = "Siete de la noche"
            b8 = "Ocho de la tarde"
            r8 = "Ocho de la noche"
            b9 = "Nueve de la tarde"
            r9 = "Nueve de la noche"
            b10 = "Diez de la tarde"
            r10 = "Diez de la noche"
            b11 = "Once de la tarde"
            r11 = "Once de la noche"
            b12 = "Doce de la mañana"
            r12 = "Doce de la noche"
            b13 = "con uno minutos"
            r13 = "con un minuto"
            b14 = "tiuno minutos"
            r14 = "tiún minutos"
            b15 = " y uno minutos,"
            r15 = " y un minutos,"
            b16 = ", uno segundos."
            r16 = ", un segundo."
            b17 = "tiuno segundos"
            r17 = "tiún segundos"
            b18 = "Uno de la "
            r18 = "Una de la "
            b19 = " con cero minutos, "
            r19 = ", "
            With doc.Sections(1).Headers(wdHeaderFooterFirstPage)
                .LinkToPrevious = True
                .range.Font.Name = fuente_s
                .range.Font.Size = fuente_s_t
                .range.ParagraphFormat.TabStops.ClearAll
                .range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), Alignment:=wdAlignTabRight
                
            End With
            
            
            s.TypeText Text:="< Hoja uno de "
            s.Fields.Add range:=s.range, Type:=wdFieldNumPages, Text:="NumPages \* Arabic \*CardText \*Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            's.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" >"
            s.TypeText Text:=vbTab
            s.Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:="."
            s.TypeParagraph
            Repeat
            s.Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date  \@ ""dd'/'MM'/'yy"" "
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=vbTab
            s.Fields.Add range:=s.range, Type:=wdFieldTime, Text:="Time \@ ""h:mm:ss am/pm"" ", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeParagraph
            s.TypeParagraph
            If el = "" Then
            GoTo 1
            End If
            If el = el Then
            GoTo 2
            End If
1           s.TypeText Text:="Carta para "
            s.TypeText Text:=Chr(171)
            s.InsertBefore (Nombre)
            s.range.Case = wdTitleWord
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=Chr(187)
            GoTo 3
2           s.TypeText Text:="Carta para"
            s.TypeText Text:=" "
            s.InsertBefore (el)
            s.range.Case = wdLowerCase
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" "
            s.TypeText Text:=Chr(171)
            s.InsertBefore (Nombre)
            s.range.Case = wdTitleWord
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=Chr(187)
3           s.ParagraphFormat.Alignment = wdAlignParagraphCenter
            s.TypeParagraph
            s.ParagraphFormat.Alignment = wdAlignParagraphLeft
            s.TypeParagraph
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="DATE \@ ""dddd"" \* FirstCap", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=", "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="DATE \@ ""d"" \* CardText \* Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" de "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="DATE \@ ""MMMM"" \* Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" de "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="DATE \@ ""yyyy"" \* CardText \* Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=". "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""h"" \* CardText", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" de la "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""am/pm"" \* CardText \* Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" con "
            s.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""mm"" \* CardText \* Lower", PreserveFormatting:=False
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" minutos, "
            s.Fields.Add range:=Selection.range, Type:=wdFieldEmpty, PreserveFormatting:=False
            s.TypeText Text:="=SUM("
            s.Fields.Add range:=Selection.range, Type:=wdFieldEmpty, Text:="TIME  \@ ""ss"" ", PreserveFormatting:=True
            s.MoveRight Unit:=wdCharacter, Count:=1
            s.TypeText Text:=",-1) \* CardText \* Lower"
            s.MoveRight Unit:=wdCharacter, Count:=1
            s.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            s.Fields.Update
            '**********No borrar esta parte entre (*)**********
            's.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""ss"" \* CardText \* Lower", PreserveFormatting:=False
            's.MoveLeft Unit:=wdWord, Count:=1, Extend:=1
            's.Fields.Update
            '***************************************************
            s.Fields.Unlink
            s.Collapse Direction:=wdCollapseEnd
            s.TypeText Text:=" segundos."
            With p.Sections(1).Headers(wdHeaderFooterFirstPage).range.Find
                .ClearFormatting
                .Text = b1
                With .Replacement
                    .ClearFormatting
                    .Text = r1
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b2
                With .Replacement
                    .ClearFormatting
                    .Text = r2
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b3
                With .Replacement
                    .ClearFormatting
                    .Text = r3
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b4
                With .Replacement
                    .ClearFormatting
                    .Text = r4
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b5
                With .Replacement
                    .ClearFormatting
                    .Text = r5
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b6
                With .Replacement
                    .ClearFormatting
                    .Text = r6
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b7
                With .Replacement
                    .ClearFormatting
                    .Text = r7
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b8
                With .Replacement
                    .ClearFormatting
                    .Text = r8
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b9
                With .Replacement
                    .ClearFormatting
                    .Text = r9
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b10
                With .Replacement
                    .ClearFormatting
                    .Text = r10
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b11
                With .Replacement
                    .ClearFormatting
                    .Text = r11
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b12
                With .Replacement
                    .ClearFormatting
                    .Text = r12
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b13
                With .Replacement
                    .ClearFormatting
                    .Text = r13
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b14
                With .Replacement
                    .ClearFormatting
                    .Text = r14
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b15
                With .Replacement
                    .ClearFormatting
                    .Text = r15
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b16
                With .Replacement
                    .ClearFormatting
                    .Text = r16
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b17
                With .Replacement
                    .ClearFormatting
                    .Text = r17
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b18
                With .Replacement
                    .ClearFormatting
                    .Text = r18
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
                .Text = b19
                With .Replacement
                    .ClearFormatting
                    .Text = r19
                End With
            .Execute Replace:=wdReplaceAll, _
                Format:=False, MatchCase:=True, _
                MatchWholeWord:=False
            End With
            s.WholeStory
            s.LanguageID = wdSpanishModernSort
            s.NoProofing = False 'Se permite revisar ortografía
    Application.CheckLanguage = True
        .SeekView = wdSeekMainDocument
    End With
    s.InsertBreak Type:=wdPageBreak
        With doc.Sections(1).Footers(wdHeaderFooterFirstPage)
            .LinkToPrevious = False
            .range.Font.Name = fuente_s
            .range.Font.Size = fuente_s_t
            .range.ParagraphFormat.TabStops.ClearAll
            .range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), Alignment:=wdAlignTabRight
        End With
    With doc.Sections(1)
        .Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
        With doc.ActiveWindow.View
            .SeekView = wdSeekFirstPageFooter
            With s
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:="/"
                .Fields.Add range:=s.range, Type:=wdFieldNumPages, Text:="NumPages \* Arabic", PreserveFormatting:=False
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=vbTab
                .TypeText Text:="A la hoja "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, PreserveFormatting:=False
                .TypeText Text:="=SUM("
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .TypeText Text:=",1)  \* CardText \* Lower"
                .Fields.Update
                .MoveRight Unit:=wdCharacter, Count:=1, Extend:=1
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" >"
            End With
            .SeekView = wdSeekMainDocument
        End With
    End With
        With doc.Sections(1).Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .range.Font.Name = fuente_s
            .range.Font.Size = fuente_s_t
            .range.ParagraphFormat.TabStops.ClearAll
            .range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), Alignment:=wdAlignTabRight
        End With
    With doc.Sections(1)
        .Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        With doc.ActiveWindow.View
            .SeekView = wdSeekPrimaryHeader
            With s
                .TypeText Text:=Chr(31) & "> de la hoja "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, PreserveFormatting:=False
                .TypeText Text:="=SUM("
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .TypeText Text:=",-1)  \* CardText \* Lower"
                .Fields.Update
                .MoveRight Unit:=wdCharacter, Count:=1, Extend:=1
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=vbTab
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:="."
            End With
        End With
    End With
        With doc.Sections(1).Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .range.Font.Name = fuente_s
            .range.Font.Size = fuente_s_t
            .range.ParagraphFormat.TabStops.ClearAll
            .range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), Alignment:=wdAlignTabRight
        End With
    With doc.Sections(1)
        .Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        With doc.ActiveWindow.View
            .SeekView = wdSeekPrimaryFooter
            With s
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:="/"
                .Fields.Add range:=s.range, Type:=wdFieldNumPages, Text:="NumPages \* Arabic", PreserveFormatting:=False
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=vbTab
                .TypeText Text:="A la hoja "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, PreserveFormatting:=False
                .TypeText Text:="=SUM("
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Arabic", PreserveFormatting:=False
                .TypeText Text:=",1)  \* CardText \* Lower"
                .Fields.Update
                .MoveRight Unit:=wdCharacter, Count:=1, Extend:=1
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" >"
            End With
            .SeekView = wdSeekMainDocument
        End With
    End With
    With s
        .TypeBackspace
        Repeat (3)
        '--- Inicia formato de párrafo ---
        With s.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .Space1
            .LineSpacing = LinesToPoints(1)
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
        End With
        '--- Finaliza formato de párrafo ---
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
'''''''''''''''''''''''''''''''''''''''''
        sangria = PointsToInches((72 / numEspaciosxPulgada)) 'En Courier tamaño 12 hay 10 espacios
        'sangria = PointsToInches(72 / numEspaciosxPulgada) 'En Courier tamaño 12 hay 10 espacios
        'sangria = PointsToInches((72 / (k / fuente_asunto_tam)) * 8) 'En Courier tamaño 12 hay 10 carácteres
        'sangria = PointsToInches((72 / (k / fuente_t)) * 8) 'En Courier tamaño 12 hay 10 carácteres
        'MsgBox PointsToInches((72 / (k / fuente_t)) * 8) '  que ocupan 72 puntos en una pulgada
        'With s.ParagraphFormat                              ' de aquí entonces el número 120 como constante
            '.LeftIndent = CentimetersToPoints(7.21)        ' sólo vale para la letra de paso constante,
          '  .FirstLineIndent = CentimetersToPoints(-2.03)  ' para las de paso variable, son valores aproximados
            '.FirstLineIndent = InchesToPoints(-sangria)     ' como la constante 180 para letra calibri tamaño 12
            '.FirstLineIndent = CentimetersToPoints(-3)
            '.FirstLineIndent = CentimetersToPoints(-1.51)
            '.LeftIndent = InchesToPoints(2.16666 + sangria)
            '.LeftIndent = InchesToPoints(2 + sangria)
            '.TabStops(Right) = InchesToPoints(2.1666)
       'End With
        
        
        With s.ParagraphFormat
        
            .TabStops.Add Position:=InchesToPoints(1), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
            .TabStops.Add Position:=InchesToPoints(3), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
            .TabStops.Add Position:=InchesToPoints(3 + sangria), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            ''.FirstLineIndent = InchesToPoints(2.1666 + 1)
            .FirstLineIndent = InchesToPoints((-3) - sangria)
            .LeftIndent = InchesToPoints((3) + sangria)
            
        End With
        
        
        
        
        
        
'''''''''''''''''''''''''''''''''''''''''
        s.Font.Size = fuente_asunto_tam
        s.Font.Name = fuente_asunto_nom
        s.TypeText Text:=vbTab
        s.TypeText Text:=vbTab
        s.TypeText Text:="Asunto:"
        s.TypeText Text:=vbTab
        s.InsertBefore (asunto)
        s.range.Case = wdTitleSentence
        s.Collapse Direction:=wdCollapseEnd
        s.Font.Name = fuente
        s.Font.Size = fuente_t
        .TypeParagraph
'''''''''''''''''''''''''''''''''''''''''
        With s.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .FirstLineIndent = CentimetersToPoints(0)
            .TabStops.ClearAll
            .TabStops.Add Position:=CentimetersToPoints(2.54)
        End With
'''''''''''''''''''''''''''''''''''''''''
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .InsertBefore (lugar)
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:=", "
        .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""d 'de' MMMM 'de' yyyy'.'"" \* Lower", PreserveFormatting:=False
        .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
        .Fields.Update
        .Fields.Unlink
        .Collapse Direction:=wdCollapseEnd
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .InsertBefore (vocativo)
        .Font.AllCaps = True
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:=" "
        .InsertBefore (Nombre)
        '.range.Case = wdUpperCase
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:=""
        .TypeParagraph
        .Font.AllCaps = False
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
'''''''''''''''''''''''''''''''''''''''''
        With s
            .ParagraphFormat.LineSpacingRule = espacio
        End With
'''''''''''''''''''''''''''''''''''''''''
        .TypeParagraph
'''''''''''''''''''''''''''''''''''''''''
        With s
            .ParagraphFormat.LineSpacingRule = espacio
            .ParagraphFormat.FirstLineIndent = InchesToPoints(0.5)
        End With
'''''''''''''''''''''''''''''''''''''''''
        .TypeParagraph
        .Collapse Direction:=wdCollapseEnd
        .TypeParagraph
        .Collapse Direction:=wdCollapseEnd
'''''''''''''''''''''''''''''''''''''''''
        With s.ParagraphFormat
            .Space1
            .FirstLineIndent = InchesToPoints(0)
        End With
'''''''''''''''''''''''''''''''''''''''''
        .TypeParagraph
        .Collapse Direction:=wdCollapseEnd
        .TypeParagraph
        .Collapse Direction:=wdCollapseEnd
        .TypeParagraph
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:="Cordialmente,"
        .TypeParagraph
        .InsertBefore (firma)
        .range.Case = wdTitleWord
        .Collapse Direction:=wdCollapseEnd
        .MoveUp Unit:=wdParagraph, Count:=8
        .InsertAfter Text:=mensajeinicial
    End With
    Application.ScreenUpdating = True
    GoTo 10
9   s.Document.Close SaveChanges:=False
10  End Sub


Sub Carta_por_correo_normal_bis_final_una_hoja()
'
' Macro56 Macro
    
    Application.ScreenUpdating = False
    Dim doc As Document
    Set doc = ActiveDocument
    Dim s As Selection
    Set s = doc.ActiveWindow.Selection
    Dim p As range
    Set p = s.Paragraphs(1).range
    Dim Nombre As String
    'Dim lineas As String
    lineas = ActiveDocument.range(WholeStory).ComputeStatistics(Statistic:=wdStatisticLines)
    'lineas = ActiveDocument.ComputeStatistics(statistic:=wdStatisticLines)
    'lineas = doc.range(WholeStory).ComputeStatistics(statistic:=wdStatisticLines)
    Dim fuente_pie As String
    fuente_pie = "Courier Final Draft"
    
    'lineas = ActiveDocument.range.ComputeStatistics(statistic:=wdStatisticLines)
    s.HomeKey Unit:=wdStory
    
    If doc.ActiveWindow.View.SplitSpecial = wdPaneNone Then
        doc.ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        doc.ActiveWindow.View.Type = wdPrintView
    End If

    With doc.ActiveWindow.View
        .SeekView = wdSeekCurrentPageHeader
            With s
                With s.Find
                    .Text = "Carta para "
                    .Replacement.Text = ""
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                s.Find.Execute
                s.MoveDown Unit:=wdParagraph, Extend:=wdExtend
                s.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                s.Copy
            End With
        .SeekView = wdSeekMainDocument
    End With

    'n = 2  '¿para qué sirve este número?, sirve para colocar el número de sección en el documento
    n = 1  '¿para qué sirve este número?, sirve para colocar el número de sección en el documento

    's.Font.Name = fuente_pie
    s.EndKey Unit:=wdStory
    
    
    
    'Esta manera ya funciona:
    'Como lo venía haciendo y funciona - cambios
    'Continuar = MsgBox("Si hay un párrafo continúo entre la ANTEPENÚLTIMA, PENÚLTIMA Y ÚLTIMA " _
              & "página u hoja de su documento pulse el botón ""Sí"", pero si no está seguro, " _
              & "clic en el botón ""Cancelar"" y revise", vbYesNoCancel, "ATENCIÓN")
    'If Continuar = vbYes Then
    '    GoTo 1
    'End If
    'If Continuar = vbNo Then
    '    GoTo 2
    'End If
    ''si añado esta parte no funciona
    'If Continuar = vbCancel Then
    '    Exit Sub
    'End If
    ' fin de cambios
'   Esta forma funciona:
    
'    Dim Mensaje, Estilo, Titulo, Respuesta
'    Mensaje = "Si hay un párrafo continúo entre la ANTEPENÚLTIMA, PENÚLTIMA Y ÚLTIMA " _
'              & "página u hoja de su documento pulse el botón ""Sí"", pero si no está seguro, " _
'              & "clic en el botón ""Cancelar"" y revise"    ' Definición del mensaje.
'    Estilo = vbYesNoCancel    ' Definición del tipo de botones.
'    Titulo = "ATENCIÓN"    ' Definición del Título.
'    Respuesta = MsgBox(Mensaje, Estilo, Titulo)   ' Definición de las respuestas.
'    If Respuesta = vbYes Then
'        GoTo 1
'    End If
'    If Respuesta = vbNo Then
'        GoTo 2
'    End If
'    If Respuesta = vbCancel Then
'        Exit Sub
'    End If
    
    
    '- cambios
    
'1   s.GoTo What:=wdGoToPage, which:=wdGoToNext, Count:=1, Name:=""
    's.InsertBreak Type:=wdSectionNewPage
    's.MoveRight unit:=wdCharacter, Count:=2
    's.MoveDown unit:=wdParagraph, Count:=1, Extend:=wdExtend
    's.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
    'GoTo 3
    'fin de cambios
    
    
    
    
    
    
    
    'cambios
    
'2   s.EndKey unit:=wdStory
 '   s.GoTo What:=wdGoToPage, which:=wdGoToNext, Count:=1, Name:=""
  '  s.MoveUp unit:=wdParagraph, Count:=1
   ' s.InsertBreak Type:=wdSectionBreakContinuous
   'fin de cambios
    
    
    
    
    's.MoveUp Unit:=wdParagraph, Count:=1 'se cambia
    's.MoveUp Unit:=wdLine, Count:=1   'SE PONE
    's.MoveRight Unit:=wdLine, Count:=1
    's.MoveLeft Unit:=wdCharacter, Count:=1
    
    's.Collapse Direction:=wdCollapseEnd 'Aquí está el movimiento al párrafo
    's.MoveLeft Unit:=wdCharacter, Count:=1
    ''s.TypeText Text:=" "
    
    ''s.MoveLeft Unit:=wdCharacter, Count:=1
    ''s.TypeText Text:=Chr(31)
    's.MoveLeft Unit:=wdCharacter, Count:=1
    ''s.MoveRight Unit:=wdCharacter, Count:=1
    's.InsertBreak Type:=wdSectionBreakContinuous
    
    
    
    
    
    
    's.MoveRight Unit:=wdCharacter, Count:=2
    's.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
    's.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)

    'doc.ParagraphFormat.LeftIndent = CentimetersToPoints(0)
    'doc.Paragraphs.First.Range.InsertBreak Type:=wdSectionBreakContinuous
    's.InsertBreak Type:=wdSectionBreakContinuous


    's.EndKey unit:=wdStory
    's.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1, Name:=""
    ''s.MoveUp unit:=wdParagraph, Count:=1 'Ojo aquí, esta dejarla sin efecto cuando n=3 con un renglón largo
    's.Collapse direction:=wdCollapseStart
    's.InsertBreak Type:=wdSectionBreakNextPage







    's.HomeKey unit:=wdStory
    's.EndKey unit:=wdStory
    's.MoveUp unit:=wdParagraph, Count:=1
    's.InsertBreak Type:=wdSectionBreakContinuous
    'doc.Paragraphs.Last.range.InsertBreak Type:=wdSectionBreakContinuous
3   s.EndKey Unit:=wdStory




    With doc.Sections(n)
        .PageSetup.DifferentFirstPageHeaderFooter = True
        '.PageSetup.DifferentFirstPageHeaderFooter = True
        '.PageSetup.BookFoldPrinting = True
        '.PageSetup.BookFoldPrinting = False
        
        '.Footers(wdHeaderFooterEvenPages).LinkToPrevious = True
    End With
    With doc.ActiveWindow.View
        .SeekView = wdSeekCurrentPageFooter
            b1 = " y uno segundos."
            r1 = " y un segundos."
            b2 = " a. m. "
            r2 = " mañana "
            b3 = " p. m. "
            r3 = " tarde "
            b4 = "Doce de la tarde"
            r4 = "Doce del día"
            b5 = "Siete de la tarde"
            r5 = "Siete de la noche"
            b6 = "Ocho de la tarde"
            r6 = "Ocho de la noche"
            b7 = "Nueve de la tarde"
            r7 = "Nueve de la noche"
            b8 = "Diez de la tarde"
            r8 = "Diez de la noche"
            b9 = "Once de la tarde"
            r9 = "Once de la noche"
            b10 = "Doce de la mañana"
            r10 = "Doce de la noche"
            b11 = "con uno minutos"
            r11 = "con un minuto"
            b12 = "tiuno minutos"
            r12 = "tiún minutos"
            b13 = " y uno minutos,"
            r13 = " y un minutos,"
            b14 = ", uno segundos."
            r14 = ", un segundo."
            b15 = "tiuno segundos"
            r15 = "tiún segundos"
            b16 = "Uno de la "
            r16 = "Una de la "
            b17 = " con cero minutos, "
            r17 = ", "
            'b18 = "y cero segundos."
            'r18 = "."
            b18 = " p.m. "
            r18 = " tarde "
            b19 = " a.m. "
            r19 = " mañana "

            With s

                .EndKey Unit:=wdStory, Extend:=wdExtend
                .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Font.AllCaps = True
                '.Font.Name = fuente_pie 'Activa esta línea si quieres cambiar el tipo de fuente del último pie de página
                .TypeText Text:="Fin de la "
            
                .PasteAndFormat (wdFormatPlainText)
            
                .TypeParagraph
                .Font.AllCaps = False
                '.ParagraphFormat.LineSpacing = LinesToPoints(0.25)
                .TypeParagraph

                '.ParagraphFormat.LineSpacing = LinesToPoints(1)
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""dddd"" \* FirstCap", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=", "
                
                '.HomeKey Unit:=wdLine
                '.Range.Case = wdTitleSentence
                '.EndKey Unit:=wdLine
                
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""d"" \* CardText \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" de "
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""MMMM"" \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" de "
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""yyyy"" \* CardText \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=". "
                '.TypeParagraph
                '.ParagraphFormat.LineSpacing = LinesToPoints(0.25)
                '.TypeParagraph
                '.ParagraphFormat.LineSpacing = LinesToPoints(1)
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""h"" \* CardText", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" de la "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""am/pm"" \* CardText \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" con "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""mm"" \* CardText \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" minutos, "
                .Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""ss"" \* CardText \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Fields.Unlink
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" segundos."
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b1
                    With .Replacement
                        .ClearFormatting
                        .Text = r1
                    End With
                    .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b2
                    With .Replacement
                        .ClearFormatting
                        .Text = r2
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b3
                    With .Replacement
                        .ClearFormatting
                        .Text = r3
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b4
                    With .Replacement
                        .ClearFormatting
                        .Text = r4
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b5
                    With .Replacement
                        .ClearFormatting
                        .Text = r5
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b6
                    With .Replacement
                        .ClearFormatting
                        .Text = r6
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b7
                    With .Replacement
                        .ClearFormatting
                        .Text = r7
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b8
                    With .Replacement
                        .ClearFormatting
                        .Text = r8
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b9
                    With .Replacement
                        .ClearFormatting
                        .Text = r9
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b10
                    With .Replacement
                        .ClearFormatting
                        .Text = r10
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b11
                    With .Replacement
                        .ClearFormatting
                        .Text = r11
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b12
                    With .Replacement
                        .ClearFormatting
                        .Text = r12
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b13
                    With .Replacement
                        .ClearFormatting
                        .Text = r13
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b14
                    With .Replacement
                        .ClearFormatting
                        .Text = r14
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b15
                    With .Replacement
                        .ClearFormatting
                        .Text = r15
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b16
                    With .Replacement
                        .ClearFormatting
                        .Text = r16
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b17
                    With .Replacement
                        .ClearFormatting
                        .Text = r17
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b18
                    With .Replacement
                        .ClearFormatting
                        .Text = r18
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                With doc.Sections(n).Footers(wdHeaderFooterFirstPage).range.Find
                    .ClearFormatting
                    .Text = b19
                    With .Replacement
                        .ClearFormatting
                        .Text = r19
                    End With
                .Execute Replace:=wdReplaceAll, _
                    Format:=False, MatchCase:=True, _
                    MatchWholeWord:=False
                End With
                
            '.SeekView = wdSeekMainDocument
        'End With
                .TypeParagraph
                .ParagraphFormat.LineSpacing = LinesToPoints(1)
                '.Fields.Add range:=s.range, Type:=wdFieldEmpty, Text:="TIME \@ ""h:mm:ss am/pm"" ", PreserveFormatting:=False
                '.MoveLeft Unit:=wdWord, count:=1, Extend:=1
                '.Fields.Update
                '.Fields.Unlink
                '.Collapse direction:=wdCollapseEnd
                '.TypeParagraph
                .TypeParagraph
                '.TypeParagraph
                .TypeText Text:="Carta con "
                .Fields.Add range:=s.range, Type:=wdFieldNumWords, Text:="NumWords", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" palabras y "
                '.TypeText Text:=vbTab
                '.TypeText Text:="Número de líneas: "
    'Dim lineas As String
    'Dim lineas As String
    'lineas = doc.range(WholeStory).ComputeStatistics(statistic:=wdStatisticLines)
                .TypeText Text:=lineas & " líneas. "

                .TypeText Text:="Se imprime el "
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""d"" \* Arabic", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" de "
                .Fields.Add range:=s.range, Type:=wdFieldDate, Text:="Date \@ ""MMMM' de 'yyyy"" \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" a las (la) "
                .Fields.Add range:=s.range, Type:=wdFieldTime, Text:="Time \@ ""h:mm:ss am/pm"" ", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=""
                .TypeParagraph

                '.TypeParagraph
                .EndKey Unit:=wdStory, Extend:=wdExtend
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                .TypeParagraph
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:="/"
                .Fields.Add range:=s.range, Type:=wdFieldNumPages, Text:="NumPages", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=vbTab
                .TypeText Text:="< Hoja "
                .Fields.Add range:=s.range, Type:=wdFieldPage, Text:="Page \* Cardtext \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" de "
                .Fields.Add range:=s.range, Type:=wdFieldNumPages, Text:="NumPages \* Cardtext \* Lower", PreserveFormatting:=False
                .MoveLeft Unit:=wdWord, Count:=1, Extend:=1
                .Fields.Update
                .Collapse Direction:=wdCollapseEnd
                .TypeText Text:=" >"
            End With
        .SeekView = wdSeekMainDocument
    End With
MsgBox ActiveDocument.ComputeStatistics(Statistic:=wdStatisticLines, _
 IncludeFootnotesAndEndnotes:=False) & " líneas"
    
10 End Sub
