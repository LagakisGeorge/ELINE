Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	Dim FILEARX As String
	Dim FILETEL As String
	Dim F_AFMEKD As String
    Dim MET(20, 6) As String
    Dim PINAKAS(20, 6) As String
    Dim PARASTATIKA(50, 4) As String
    Dim fn_parastat As Integer ' megeuow PARASTATIKA
	
	
	
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim A As String
		Dim B, c As Object
        'metatrepei eggrafo apo 437->9
		's928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
		's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι" ' saehioyv
		's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"
		' saehioyv
		'If Len(Dir(Text1.Text)) < 2 Then
		'  MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ " + Text1.Text
		' Exit Sub
		'End If
		
		
		FileOpen(2, Text2.Text, OpenMode.Output)
		
		FileOpen(1, Text1.Text, OpenMode.Input)
		Do While Not EOF(1)
			A = LineInput(1)
			' ΔΙΑΒΑΖΩ ΤΑ ΣΤΟΙΧΕΙΑ ΠΟΥ ΘΕΛΩ ΓΙΑ ΝΑ ΔΗΜΙΟΥΡΓΗΣΩ ΤΗΝ Ε_LINE
			
			
			
			
			
			
			
			PrintLine(2, A)
		Loop 
		FileClose(1)
		
		' pRINT#2, E_LINE
		
		FileClose(2)
		
		
		Kill(FILEARX)
		'Shell "c:\NOTEPAD.EXE /P C:\LAGEURO\FORMA2.TXT", vbMaximizedFocus
		
		
		
		
		
		
		' MsgBox "ok"
		
		
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim K As Object
		'=====================================================================================
		Dim A As String
		Dim B, c As Object
		'metatrepei eggrafo apo 437->928
		's928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
		's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι" ' saehioyv
		's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"
		' saehioyv
		'If Len(Dir(Text1.Text)) < 2 Then
		'  MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ " + Text1.Text
		' Exit Sub
		'End If
		Dim ff As String
		
        Dim m_pros As Integer

        m_pros = 1
		
		Dim e As String
		'[<]099368090;099368090;12345678;070420131200;221;A;1001;0.00;3.00;3.00;0.00;0.00;0.00;0.39;0.69;0.00;7.08;0[>]
		e = "[<]" & F_AFMEKD
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		ff = Dir(Text1.Text & "\*.PRN", FileAttribute.Normal)
		
		
		ff = Text1.Text & "\" & ff
		
		For K = 1 To 17
			'Debug.Print Str(K), PINAKAS(K, 0), PINAKAS(K, 1), PINAKAS(K, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PINAKAS(K, 4) = ""
		Next 
		
		
		
		'Open Text2.Text For Output As #2
		Dim N As Short
		N = 0
        Dim l As Integer
		
        Dim st As Integer
        Dim c2 As String
        Dim c3, c4 As String
        load_pinakas(1)
        '
		FileOpen(1, ff, OpenMode.Input)
		Do While Not EOF(1)
			A = LineInput(1)
			' ΔΙΑΒΑΖΩ ΤΑ ΣΤΟΙΧΕΙΑ ΠΟΥ ΘΕΛΩ ΓΙΑ ΝΑ ΔΗΜΙΟΥΡΓΗΣΩ ΤΗΝ Ε_LINE
            N = N + 1

            'ΕΛΑΓΧΟΣ ΑΝ ΕΙΝΑΙ Η ΦΟΡΜΑ 1 Η Η ΦΟΡΜΑ 2
            If N = 1 Then
                'ΑΝ ΒΡΙΣΚΕΙ ΤΗΝ ΗΜΕΡΟΜΗΝΙΑ ΣΤΗΝ ΣΩΣΤΗ ΘΕΣΗ ΣΗΜΑΙΝΕΙ ΟΚ
                ' ΑΛΛΟΙΩΣ ΠΑΩ 2Η ΦΟΡΜΑ
                If Mid(A, PINAKAS(3, 2) + 2, 1) = "/" Then
                    load_pinakas(1)
                Else
                    load_pinakas(2)
                End If
            End If





            'st = Val(PINAKAS(K, 2))
            'If st = 0 Then st = 1
            'c = Mid(A, st, Val(PINAKAS(K, 3)))
            If N = 45 Then
                c2 = Mid(A, 68, 10)
                If Val(Replace(c2, ",", ".")) > 0 Then
                    'ok  sthn 49 seira to synolo
                Else
                    'προσθετο +1 για τα σύνολα 
                    For K = 7 To 16
                        PINAKAS(K, 1) = PINAKAS(K, 1) + 1
                    Next

                End If

            End If






            Dim nk, nt As Integer
            Dim mpros2 As Integer

            'ΨΑΧΝΩ ΝΑ ΒΡΩ ΠΟΙΑ ΑΝΑΦΕΡΟΝΤΑΙ ΣΕ ΑΥΤΗΝ ΤΗΝ ΣΕΙΡΑ
            For K = 1 To 17
                If Val(PINAKAS(K, 1)) = N Then
                    c = Mid(A, Val(PINAKAS(K, 2)), Val(PINAKAS(K, 3)))
                    mpros2 = 1
                    If K >= 7 And K <= 16 Then
                        nt = InStr(c, ".")
                        nk = InStr(c, ",")
                        If nt > 0 And nk > 0 Then ' εχει και κομα και τελεια ο αριθμός
                            If nk > nt Then 'για δεκαδικο εχει το κομα
                                c = Replace(c, ".", "")
                                c = Replace(c, ",", ".")
                            Else
                                'για δεκαδικό εχει την τελεία
                                c = Replace(c, ",", "")  ' εξαφανίζω τα κόμματα
                            End If
                        End If
                        If InStr(c, "-") > 0 Then
                            mpros2 = -1  'σημαινει οτι ο αριθμος είναι αρνητικός
                            c = Replace(c, "-", "") ' σβηνω το - για να μην με μπερδευει

                        End If
                        PINAKAS(K, 4) = VB6.Format(m_pros * mpros2 * Val(Replace(c, ",", ".")), "#######.00")
                        'End If
                    ElseIf K = 5 Then ' αν ειναι ο τιτλος του παραστατικού

                        Me.Text = c
                        'PARASTATIKA(l, n)  n:  0=kod	1=titlos	2=pros	3=use
                        For l = 1 To fn_parastat

                            c3 = Mid(LTrim(c), 1, 15)
                            c4 = Mid((PARASTATIKA(l, 1)), 1, 15)
                            If InStr(c3, c4) > 0 Then
                                PINAKAS(K, 4) = PARASTATIKA(l, 0) ' "222"
                                m_pros = PARASTATIKA(l, 2)
                            End If
                        Next                        'If InStr(c, Mid("ƒ. €§¦©«¦??? - ’ £¦??? ¦", 1, 5)) > 0 Then
                        '    PINAKAS(K, 4) = "222" 'TIM.POL
                        'ElseIf InStr(c, "ΛΙΑΝ") > 0 Then
                        '    PINAKAS(K, 4) = "231" 'LIANIKH
                        'Else
                        '    PINAKAS(K, 4) = c

                        'End If


                    Else


                        PINAKAS(K, 4) = LTrim(Replace(c, ".", ""))  ' c
                    End If

                End If
                ' Debug.Print PINAKAS(K, 0), K   ',pinakas(k,1)
            Next
        Loop
		FileClose(1)
		
        'Print#(2, E_LINE)
        'Type Visual Basic 6 code here...
        'Print#(2, E_LINE)
        'UPGRADE_ISSUE: The preceding line couldn't be parsed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"'

		
        Dim M_HME As String
        Dim nn1, nn2 As Integer, mday As String, mmon As String, myear As String
        nn1 = InStr(PINAKAS(3, 4), "/")
        nn2 = InStrRev(PINAKAS(3, 4), "/")


        If nn1 = 3 Then
            mday = Mid(PINAKAS(3, 4), 1, 2)
            If nn2 = 5 Then
                mmon = "0" + Mid(PINAKAS(3, 4), 4, 1)
            Else
                mmon = Mid(PINAKAS(3, 4), 4, 2)
            End If
        Else  'nn1=2
            mday = "0" + Mid(PINAKAS(3, 4), 1, 1)

            If nn2 = 4 Then
                mmon = "0" + Mid(PINAKAS(3, 4), 3, 1)
            Else  ' 5
                mmon = Mid(PINAKAS(3, 4), 3, 2)
            End If
        End If
        myear = Mid(PINAKAS(3, 4), nn2 + 1, 4)
        If Val(myear) > 2000 Then
            'ok
        Else
            myear = "20" + Mid(PINAKAS(3, 4), nn2 + 1, 2)
        End If






        Dim mores As String = Mid(PINAKAS(17, 4), 1, 5)

        nn1 = InStr(mores, ":")
        If nn1 = 2 Then '2:35
            mores = "0" + Mid(mores, 1, 1) + Mid(mores, 3, 2)
        Else
            mores = Mid(mores, 1, 2) + Mid(mores, 4, 2)
        End If
        

        '12/11/15H G10:09
        M_HME = myear & mmon & mday & mores

        'PINAKAS(2, 4) &
        e = e & ";" & PINAKAS(1, 4) & ";;" & M_HME  'ΑΦΜ,,ΚΑΡΤΑ,ΗΜΕΡΟΜΗΝΙΑ, , , , (3KENA)

        e = e & ";" & PINAKAS(5, 4) & ";" & PINAKAS(4, 4) & ";" & PINAKAS(6, 4)   ' ειδοσ 222 , σειρα , αριθμος


        'ΑΝ ΚΑΘ13+Καθ24+καθ6 =0 τοτε  καθ0=τελικο
        If Val(Replace(PINAKAS(7, 4), ",", ".")) + Val(Replace(PINAKAS(8, 4), ",", ".")) + Val(Replace(PINAKAS(9, 4), ",", ".")) = 0 Then

            PINAKAS(11, 4) = PINAKAS(16, 4)


        End If


        For K = 7 To 16
            'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            e = e & ";" & Replace(VB6.Format(Val(Replace(PINAKAS(K, 4), ",", ".")), "#####0.00"), ",", ".")
        Next
        e = e & ";0[>]"  'euro

        FileOpen(1, ff, OpenMode.Append)
        WriteLine(1, e)
        FileClose(1)



        FileCopy(ff, FILETEL)
        MILSEC(1000)

        Kill(ff)  'FILEARX
        ''







    End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Cd1Open.ShowDialog()
		Text1.Text = Cd1Open.FileName
		
		
		
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(FILEARX & "\*.PRN", FileAttribute.Normal)) > 0 Then
			
			Command2_Click(Command2, New System.EventArgs())
			MILSEC(5000)
			
			
		End If
		
		
		
		
	End Sub
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim L As Object
		Dim A As String
		Dim TEMP() As String
		Dim K As Short
		K = 0
		Dim MESS As String
		
        On Error GoTo ff
		FileOpen(3, My.Application.Info.DirectoryPath & "\FILES.TXT", OpenMode.Input) 'ΑΡΧΙΚΟ ΦΙΛΕ ΚΑΙ ΤΕΛΙΚΟ
		FILEARX = LineInput(3)
		MESS = "FILETEL 2H ΣΕΙΡΑ ΕΧΕΙ ΛΑΘΟΣ"
		FILETEL = LineInput(3)
		MESS = "ΑΦΜ ΕΚΔΟΤΗ 3H ΣΕΙΡΑ ΕΧΕΙ ΛΑΘΟΣ"
		F_AFMEKD = LineInput(3)
		For K = 1 To 17
			MESS = "ΣΤΟΙΧΕΙΑ ΣΕΙΡΑ ,ΣΤΗΛΗ ΣΤΗΝ " & Str(K + 3) & " ΣΕΙΡΑ ΕΧΕΙ ΛΑΘΟΣ"
			A = LineInput(3)
			
			TEMP = Split(A, ";")
			MET(K, 0) = TEMP(0)
			MET(K, 1) = TEMP(1)
			MET(K, 2) = TEMP(2)
            MET(K, 3) = TEMP(3)

            MET(K, 4) = TEMP(4)
            MET(K, 5) = TEMP(5)
			
		Next 
		
		FileClose(3)
		Text1.Text = FILEARX
		Text2.Text = FILETEL
		
		
        load_pinakas(1)
        load_parastatika()
		
        'For K = 1 To 17
        '	Debug.Print(VB6.TabLayout(Str(K), PINAKAS(K, 0), PINAKAS(K, 1), PINAKAS(K, 2)))
        '	PINAKAS(K, 4) = ""
        'Next 
		
		
		Exit Sub
		
ff: 
		MsgBox(MESS)
		Resume Next
		
		
    End Sub

    Sub load_parastatika()
        Dim c() As String
        Dim b As String = My.Application.Info.DirectoryPath & "\PARASTATIKAsyn.CSV"
        Dim cc As String
        FileOpen(1, b, OpenMode.Input) 'ΑΡΧΙΚΟ ΦΙΛΕ ΚΑΙ ΤΕΛΙΚΟ
        Dim n As Integer
        Do While Not EOF(1)
            n = n + 1
            cc = LineInput(1)
            c = Split(cc, ";")
            PARASTATIKA(n, 0) = c(0) '0=kod	1=titlos	2=pros	3=use
            PARASTATIKA(n, 1) = c(1)
            PARASTATIKA(n, 2) = c(2)
            PARASTATIKA(n, 3) = c(3)
        Loop

        fn_parastat = n

        FileClose(1)


    End Sub


    Sub load_pinakas(ByVal FORMA As Integer)
        Dim B As String

        Dim k As Integer

        Dim l As Integer
        For k = 1 To 17
            B = MET(k, 0)
            '   Debug.Print B
            'Select Case Β
            If B = "AFM" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(1, l) = MET(k, l) : Next
                'PINAKAS(1, 3) = 9 'mhkos
            ElseIf B = "KAR" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(2, l) = MET(k, l) : Next
                'PINAKAS(2, 3) = 9 'mhkos
            ElseIf B = "HME" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(3, l) = MET(k, l) : Next
                'PINAKAS(3, 3) = 9 'mhkos
            ElseIf B = "SEI" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(4, l) = MET(k, l) : Next
                'PINAKAS(4, 3) = 9 'mhkos

            ElseIf B = "EID" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(5, l) = MET(k, l) : Next
                'PINAKAS(5, 3) = 9 'mhkos



            ElseIf B = "ARI" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(6, l) = MET(k, l) : Next
                'PINAKAS(6, 3) = 9 'mhkos
            ElseIf B = "KA1" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(7, l) = MET(k, l) : Next
                'PINAKAS(7, 3) = 9 'mhkos
            ElseIf B = "KA2" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(8, l) = MET(k, l) : Next
                'PINAKAS(8, 3) = 9 'mhkos
            ElseIf B = "KA3" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(9, l) = MET(k, l) : Next
                'PINAKAS(9, 3) = 9 'mhkos
            ElseIf B = "KA4" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(10, l) = MET(k, l) : Next
                'PINAKAS(10, 3) = 9 'mhkos


            ElseIf B = "KA5" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(11, l) = MET(k, l) : Next
                'PINAKAS(11, 3) = 9 'mhkos
            ElseIf B = "FP1" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(12, l) = MET(k, l) : Next
                'PINAKAS(12, 3) = 9 'mhkos
            ElseIf B = "FP2" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(13, l) = MET(k, l) : Next
                'PINAKAS(13, 3) = 9 'mhkos
            ElseIf B = "FP3" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(14, l) = MET(k, l) : Next
                'PINAKAS(14, 3) = 9 'mhkos
            ElseIf B = "FP4" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(15, l) = MET(k, l) : Next


            ElseIf B = "SYN" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(16, l) = MET(k, l) : Next


            ElseIf B = "ORA" Then
                For l = 0 To 5
                    'UPGRADE_WARNING: Couldn't resolve default property of object L. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PINAKAS(17, l) = MET(k, l) : Next



            End If









        Next

        If FORMA = 2 Then
            For k = 1 To 17
                PINAKAS(k, 1) = PINAKAS(k, 4)
                PINAKAS(k, 2) = PINAKAS(k, 5)
            Next

        End If



    End Sub



End Class