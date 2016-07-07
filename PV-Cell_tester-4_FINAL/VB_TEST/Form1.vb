Imports System.IO
Imports System.Collections
Imports System.Diagnostics
Imports System.Data.DataTable
Imports System.Runtime.InteropServices
Public Class Form11 ' private member variables
    Public table As New DataTable("Table")
    Private mFolder As String
    Private mImageList As ArrayList
    Private mImagePosition As Integer
    Private mPreviousImage As Image
    Private mCurrentImage As Image
    Private mNextImage As Image
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public f As FileInfo
    Public xx As Form_stat = New Form_stat
    Public num As Integer = 1
    Public otp_a As Double  ' otp detect a sub cell
    Public otp_b As Double
    Public otp_c As Double
    Public otp_aG As Double  ' otp detect a sub cell
    Public otp_bG As Double
    Public otp_cG As Double
    Public otp As Double = 0
    Public otp_otp As Double = 0
    Public otp_gli As Double = 0
    Public otp1 As Double = 0
    Public gld As Double = 0
    Public gld1 As Double = 0


    ''''''''AAAAAAA
    Public z As Double = 0 ' Dark
    Public x As Double = 0 ' Dim
    Public c As Double = 0 ' L-Grey
    Public v As Double = 0 ' Light
    Public b As Double = 0 ' Sm▲
    Public n As Double = 0 ' Sm▼
    Public m As Double = 0 ' Big▼
    Public a As Double = 0 ' Big▲
    Public s As Double = 0 ' ↔
    Public d As Double = 0 ' Gradi_error

    '''''''' BBBBB
    Public zB As Double = 0 ' Dark
    Public xB As Double = 0 ' Dim
    Public cB As Double = 0 ' L-Grey
    Public vB As Double = 0 ' Light
    Public bB As Double = 0 ' Sm▲
    Public nB As Double = 0 ' Sm▼
    Public mB As Double = 0 ' Big▼
    Public aB As Double = 0 ' Big▲
    Public sB As Double = 0 ' ↔
    Public dB As Double = 0 ' Gradi_error

    ''''''''CCCCCC

    Public zC As Double = 0 ' Dark
    Public xC As Double = 0 ' Dim
    Public cC As Double = 0 ' L-Grey
    Public vC As Double = 0 ' Light
    Public bC As Double = 0 ' Sm▲
    Public nC As Double = 0 ' Sm▼
    Public mC As Double = 0 ' Big▼
    Public aC As Double = 0 ' Big▲
    Public sC As Double = 0 ' ↔
    Public dC As Double = 0 ' Gradi_error


    Public z1 As Double = 0 ' Dark
    Public x1 As Double = 0 ' Dim
    Public c1 As Double = 0 ' L-Grey
    Public v1 As Double = 0 ' Light
    Public b1 As Double = 0 ' Sm▲
    Public n1 As Double = 0 ' Sm▼
    Public m1 As Double = 0 ' Big▼
    Public a1 As Double = 0 ' Big▲
    Public s1 As Double = 0 ' ↔
    Public d1 As Double = 0 ' Gradi_error


    Public dd As Double = 0
    Public dd2 As Double = 0
    Public xxx As Integer = 1
    <DllImport("user32.dll")> Public Shared Function SetParent(ByVal hwndChild As IntPtr, ByVal hwndNewParent As IntPtr) As Integer
    End Function
    Private Sub Form11_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        For count As Integer = My.Application.OpenForms.Count - 1 To 1 Step -1
            My.Application.OpenForms(count).Close()
        Next
    End Sub
    Private Sub Add_Row_To_DataGridView_Using_TextBoxes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        table.Columns.Add("#", Type.GetType("System.Int32"))
        table.Columns.Add("CELL Id", Type.GetType("System.Int32"))
        table.Columns.Add("A_zone-HORIZ", Type.GetType("System.String"))
        table.Columns.Add("B_zone-HORIZ", Type.GetType("System.String"))
        table.Columns.Add("C_zone-HORIZ", Type.GetType("System.String"))
        table.Columns.Add("A_zone-VERT", Type.GetType("System.String"))
        table.Columns.Add("B_zone-VERT", Type.GetType("System.String"))
        table.Columns.Add("C_zone-VERT", Type.GetType("System.String"))
        table.Columns.Add("OTP - ❶", Type.GetType("System.String"))
        table.Columns.Add("GLI - ❷", Type.GetType("System.String"))
        DataGridView1.DataSource = table
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).Width = 90
        DataGridView1.Columns(3).Width = 90
        DataGridView1.Columns(4).Width = 90
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).Width = 90
        DataGridView1.Columns(8).Width = 85
        DataGridView1.Columns(9).Width = 83
    End Sub
    Private Sub MAIN_CALC()
        Dim Res As Double() = {}
        Dim Res2 As Double() = {}
        Dim Res3 As Double() = {}
        Dim Res4 As Double() = {}
        Dim Res5 As Double() = {}
        Dim Res6 As Double() = {}

        z = 0 ' Dark
        x = 0 ' Dim
        c = 0 ' L-Grey
        v = 0 ' Light
        b = 0 ' Sm▲
        n = 0 ' Sm▼
        m = 0 ' Big▼
        a = 0 ' Big▲
        s = 0 ' ↔
        d = 0 ' Gradi_error

        '''''''''BBBBB
        zB = 0 ' Dark
        xB = 0 ' Dim
        cB = 0 ' L-Grey
        vB = 0 ' Light
        bB = 0 ' Sm▲
        nB = 0 ' Sm▼
        mB = 0 ' Big▼
        aB = 0 ' Big▲
        sB = 0 ' ↔
        dB = 0 ' Gradi_error


        ''''''''CCCCC
        zC = 0 ' Dark
        xC = 0 ' Dim
        cC = 0 ' L-Grey
        vC = 0 ' Light
        bC = 0 ' Sm▲
        nC = 0 ' Sm▼
        mC = 0 ' Big▼
        aC = 0 ' Big▲
        sC = 0 ' ↔
        dC = 0 ' Gradi_error

        dd = 0 'su
        otp = 0
        otp1 = 0
        gld = 0
        gld1 = 0
        Dim TestString As String = (A_top_text.Text)
        Dim TestArray() As String = TestString.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString2 As String = (B_top_text.Text)
        Dim TestArray2() As String = TestString2.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString3 As String = (C_top_text.Text)
        Dim TestArray3() As String = TestString3.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString4 As String = (A_butt_text.Text)
        Dim TestArray4() As String = TestString4.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString5 As String = (B_butt_text.Text)
        Dim TestArray5() As String = TestString5.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString6 As String = (C_butt_text.Text)
        Dim TestArray6() As String = TestString6.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim TestString_CELL As String = (TextBox_CELL_NUM.Text)
        Dim TestArray_CELL() As String = TestString_CELL.Split(New String() {Environment.NewLine},
                                       StringSplitOptions.None)

        Dim LastNonEmpty As Integer = -1
        RichText_RES.Text = ""
        RichText_RES2.Text = ""
        Dim Rang_GRAD As Integer  '''''GRADI RANGE 
        Rang_GRAD = Val(Rang_GRADI.Text)

        Dim Rang_GSC_ As Integer  '''''RANGE FOR Grey Scale level
        Rang_GSC_ = Val(Rang_GSC.Text)

        Dim B_D_L_ As Integer ''''Bottom DARK LEVEL 
        B_D_L_ = Val(B_D_L.Text)

        Dim T_DIM_L_ As Integer
        T_DIM_L_ = Val(T_DIM_L.Text)

        Dim B_Light_L_ As Integer
        B_Light_L_ = Val(B_Light_L.Text)

        Dim U_N_F_lvl_ As Integer
        U_N_F_lvl_ = Val(U_N_F_lvl.Text)

        Dim j As Integer = 0
        Dim j1 As Integer = 1
        Dim ax, adx1, adx2, adx3, adx4 As Integer
        num = 1
        '''''''''''''TABLE COLUMNS SIZE
        table.Clear()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).Width = 90
        DataGridView1.Columns(3).Width = 90
        DataGridView1.Columns(4).Width = 90
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).Width = 90
        DataGridView1.Columns(8).Width = 76
        DataGridView1.Columns(9).Width = 76

        For i As Integer = 0 To TestArray.Length - 1
            otp_a = 0
            otp_b = 0
            otp_c = 0
            otp_aG = 0
            otp_bG = 0
            otp_cG = 0

            If TestArray(i) <> "" Then
                LastNonEmpty += 1

                TestArray(LastNonEmpty) = TestArray(i)      ' A_top_text.Text
                TestArray2(LastNonEmpty) = TestArray2(i)    ' B_top_text.Text
                TestArray3(LastNonEmpty) = TestArray3(i)    ' C_top_text.Text
                TestArray4(LastNonEmpty) = TestArray4(i)    ' A_butt_text.Text
                TestArray5(LastNonEmpty) = TestArray5(i)    ' B_butt_text.Text
                TestArray6(LastNonEmpty) = TestArray6(i)    ' C_butt_text.Text
                TestArray_CELL(LastNonEmpty) = TestArray_CELL(i)

                ''''''NUMBER of cell
                RichText_RES.Text += num & " - " & TestArray_CELL(j) & vbTab
                RichText_RES2.Text += num & " - " & TestArray_CELL(j) '& vbTab

                Array.Resize(Res, Res.Length + 1)
                Res(Res.Length - 1) = Convert.ToDouble(TestArray(i))        ' A_top_text.Text

                Array.Resize(Res2, Res2.Length + 1)
                Res2(Res2.Length - 1) = Convert.ToDouble(TestArray2(i))     ' B_top_text.Text

                Array.Resize(Res3, Res3.Length + 1)
                Res3(Res3.Length - 1) = Convert.ToDouble(TestArray3(i))     ' C_top_text.Text

                Array.Resize(Res4, Res4.Length + 1)
                Res4(Res4.Length - 1) = Convert.ToDouble(TestArray4(i))     ' A_butt_text.Text

                Array.Resize(Res5, Res5.Length + 1)
                Res5(Res5.Length - 1) = Convert.ToDouble(TestArray5(i))     ' B_butt_text.Text

                Array.Resize(Res6, Res6.Length + 1)
                Res6(Res6.Length - 1) = Convert.ToDouble(TestArray6(i))     ' C_butt_text.Text

                With ProgressBar1
                    .Style = ProgressBarStyle.Blocks
                    .Step = 0.1
                    .Minimum = 0
                    .Maximum = i
                    .Value = i
                End With
                ''''''''''''''''''''''''''''Conditions''''''''''''''''''''''''''''
                Dim __dark As String = " DRK "
                Dim __dim As String = " DIM "
                Dim __lgry As String = "LGry"
                Dim __lght As String = "LGHT"
                Dim __none As String = "####"
                Dim __lvl As String = "↑lvl↓"
                Dim __DSm As String = "↓Sml"
                Dim __TSm As String = "↑Sml"
                Dim __Dbig As String = "↓Big"
                Dim __TBig As String = "↑Big"
                ''''''''''''''''''''''''''''A-DARK''''''''''''''''''''''''''
                adx1 = 0
                adx2 = 0
                adx3 = 0
                adx4 = 0
                If Res(Res.Length - 1) <= B_D_L_ Then adx1 = 1
                If (Res(Res.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res(Res.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx2 = 1
                If Res4(Res4.Length - 1) <= B_D_L_ Then adx3 = 1
                If (Res4(Res4.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res4(Res4.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx4 = 1
                If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                    RichText_RES.Text += " - " & __dark & vbTab
                    z = z + 1
                    otp_a = 0
                    '''''''''''''''''''''''''''''''''''''''A-DARK-END''''
                    '''''''''''''''''''''''''''''''''''''''DIM''''
                Else
                    adx1 = 0
                    adx2 = 0
                    adx3 = 0
                    adx4 = 0
                    If Res(Res.Length - 1) > B_D_L_ And Res(Res.Length - 1) <= T_DIM_L_ Then adx1 = 1
                    If (Res(Res.Length - 1) + Rang_GSC_) > B_D_L_ And (Res(Res.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res(Res.Length - 1) - Rang_GSC_) > B_D_L_ And (Res(Res.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx2 = 1
                    If Res4(Res4.Length - 1) > B_D_L_ And Res4(Res4.Length - 1) <= T_DIM_L_ Then adx3 = 1
                    If (Res4(Res4.Length - 1) + Rang_GSC_) > B_D_L_ And (Res4(Res4.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res4(Res4.Length - 1) - Rang_GSC_) > B_D_L_ And (Res4(Res4.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx4 = 1

                    If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                        RichText_RES.Text += " - " & __dim & vbTab
                        x = x + 1
                        otp_a = 1
                        '''''''''''''''''''''''''''''''''''''''A-DIM-END''''
                        '''''''''''''''''''''''''''''''''''''''A-Light_GREY''''
                    Else
                        adx1 = 0
                        adx2 = 0
                        adx3 = 0
                        adx4 = 0
                        If Res(Res.Length - 1) > T_DIM_L_ And Res(Res.Length - 1) <= B_Light_L_ Then adx1 = 1
                        If (Res(Res.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res(Res.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res(Res.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res(Res.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx2 = 1
                        If Res4(Res4.Length - 1) > T_DIM_L_ And Res4(Res4.Length - 1) <= B_Light_L_ Then adx3 = 1
                        If (Res4(Res4.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res4(Res4.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res4(Res4.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res4(Res4.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx4 = 1

                        If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                            RichText_RES.Text += " - " & __lgry & vbTab
                            c = c + 1
                            otp_a = 2
                            '''''''''''''''''''''''''''''''''''''''A-Light_GREY-END''''
                            '''''''''''''''''''''''''''''''''''''''A-Light''''
                        Else
                            adx1 = 0
                            adx2 = 0
                            adx3 = 0
                            adx4 = 0
                            If Res(Res.Length - 1) > B_Light_L_ Then adx1 = 1
                            If Res4(Res4.Length - 1) > B_Light_L_ Then adx3 = 1
                            If (Res(Res.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx2 = 1
                            If (Res4(Res4.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx4 = 1

                            If adx1 = 1 And adx3 = 1 Or adx1 = 1 And adx3 = 1 And adx2 = 1 And adx4 = 1 Then
                                RichText_RES.Text += " - " & __lght & vbTab
                                v = v + 1
                                otp_a = 3
                            Else
                                RichText_RES.Text += " - " & __none & vbTab
                                d = d + 1
                                otp_a = 4
                            End If
                        End If
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''A-Light-END''''
                '''''''''''''''''''''''''''''''''''''''A-END
                '
                '''''''''''''''''''''''''''''''''''''''B
                '''''''''''''''''''''''''''''''''''''''B-DARK''''
                adx1 = 0
                adx2 = 0
                adx3 = 0
                adx4 = 0
                If Res2(Res2.Length - 1) <= B_D_L_ Then adx1 = 1
                If (Res2(Res2.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res2(Res2.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx2 = 1
                If Res5(Res5.Length - 1) <= B_D_L_ Then adx3 = 1
                If (Res5(Res5.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res5(Res5.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx4 = 1
                If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                    RichText_RES.Text += " - " & __dark & vbTab
                    zB = zB + 1
                    otp_b = 0
                    '''''''''''''''''''''''''''''''''''''''B-DARK-END''''
                    '''''''''''''''''''''''''''''''''''''''DIM''''
                Else
                    adx1 = 0
                    adx2 = 0
                    adx3 = 0
                    adx4 = 0
                    If Res2(Res2.Length - 1) > B_D_L_ And Res2(Res2.Length - 1) <= T_DIM_L_ Then adx1 = 1
                    If (Res2(Res2.Length - 1) + Rang_GSC_) > B_D_L_ And (Res2(Res2.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res2(Res2.Length - 1) - Rang_GSC_) > B_D_L_ And (Res2(Res2.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx2 = 1
                    If Res5(Res5.Length - 1) > B_D_L_ And Res5(Res5.Length - 1) <= T_DIM_L_ Then adx3 = 1
                    If (Res5(Res5.Length - 1) + Rang_GSC_) > B_D_L_ And (Res5(Res5.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res5(Res5.Length - 1) - Rang_GSC_) > B_D_L_ And (Res5(Res5.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx4 = 1

                    If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                        RichText_RES.Text += " - " & __dim & vbTab
                        xB = xB + 1
                        otp_b = 1
                        '''''''''''''''''''''''''''''''''''''''B-DIM-END''''
                        '''''''''''''''''''''''''''''''''''''''B-Light_GREY''''
                    Else
                        adx1 = 0
                        adx2 = 0
                        adx3 = 0
                        adx4 = 0
                        If Res2(Res2.Length - 1) > T_DIM_L_ And Res2(Res2.Length - 1) <= B_Light_L_ Then adx1 = 1
                        If (Res2(Res2.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res2(Res2.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res2(Res2.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res2(Res2.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx2 = 1
                        If Res5(Res5.Length - 1) > T_DIM_L_ And Res5(Res5.Length - 1) <= B_Light_L_ Then adx3 = 1
                        If (Res5(Res5.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res5(Res5.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res5(Res5.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res5(Res5.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx4 = 1

                        If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                            RichText_RES.Text += " - " & __lgry & vbTab
                            cB = cB + 1
                            otp_b = 2
                            '''''''''''''''''''''''''''''''''''''''B-Light_GREY-END''''
                            '''''''''''''''''''''''''''''''''''''''B-Light''''
                        Else

                            adx1 = 0
                            adx2 = 0
                            adx3 = 0
                            adx4 = 0
                            If Res2(Res2.Length - 1) > B_Light_L_ Then adx1 = 1
                            If Res5(Res5.Length - 1) > B_Light_L_ Then adx3 = 1
                            If (Res2(Res2.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx2 = 1
                            If (Res5(Res5.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx4 = 1

                            If adx1 = 1 And adx3 = 1 Or adx1 = 1 And adx3 = 1 And adx2 = 1 And adx4 = 1 Then
                                RichText_RES.Text += " - " & __lght & vbTab
                                vB = vB + 1
                                otp_b = 3

                            Else
                                RichText_RES.Text += " - " & __none & vbTab
                                dB = dB + 1
                                otp_b = 4

                            End If
                        End If
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''B-Light-END''''
                '''''''''''''''''''''''''''''''''''''''B-END
                '
                '''''''''''''''''''''''''''''''''''''''C
                '''''''''''''''''''''''''''''''''''''''C-DARK''''
                adx1 = 0
                adx2 = 0
                adx3 = 0
                adx4 = 0
                If Res3(Res3.Length - 1) <= B_D_L_ Then adx1 = 1
                If (Res3(Res3.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res3(Res3.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx2 = 1
                If Res6(Res6.Length - 1) <= B_D_L_ Then adx3 = 1
                If (Res6(Res6.Length - 1) + Rang_GSC_) <= B_D_L_ Or (Res6(Res6.Length - 1) - Rang_GSC_) <= B_D_L_ Then adx4 = 1
                If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                    RichText_RES.Text += " - " & __dark '& vbCrLf
                    zC = zC + 1
                    otp_c = 0
                    '''''''''''''''''''''''''''''''''''''''C-DARK-END''''
                    '''''''''''''''''''''''''''''''''''''''DIM''''
                Else
                    adx1 = 0
                    adx2 = 0
                    adx3 = 0
                    adx4 = 0
                    If Res3(Res3.Length - 1) > B_D_L_ And Res3(Res3.Length - 1) <= T_DIM_L_ Then adx1 = 1
                    If (Res3(Res3.Length - 1) + Rang_GSC_) > B_D_L_ And (Res3(Res3.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res3(Res3.Length - 1) - Rang_GSC_) > B_D_L_ And (Res3(Res3.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx2 = 1
                    If Res6(Res6.Length - 1) > B_D_L_ And Res6(Res6.Length - 1) <= T_DIM_L_ Then adx3 = 1
                    If (Res6(Res6.Length - 1) + Rang_GSC_) > B_D_L_ And (Res6(Res6.Length - 1) + Rang_GSC_) <= T_DIM_L_ Or (Res6(Res6.Length - 1) - Rang_GSC_) > B_D_L_ And (Res6(Res6.Length - 1) - Rang_GSC_) <= T_DIM_L_ Then adx4 = 1

                    If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                        RichText_RES.Text += " - " & __dim '& vbCrLf
                        xC = xC + 1
                        otp_c = 1
                        '''''''''''''''''''''''''''''''''''''''C-DIM-END''''
                        '''''''''''''''''''''''''''''''''''''''C-Light_GREY''''
                    Else
                        adx1 = 0
                        adx2 = 0
                        adx3 = 0
                        adx4 = 0
                        If Res3(Res3.Length - 1) > T_DIM_L_ And Res3(Res3.Length - 1) <= B_Light_L_ Then adx1 = 1
                        If (Res3(Res3.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res3(Res3.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res3(Res3.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res3(Res3.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx2 = 1
                        If Res6(Res6.Length - 1) > T_DIM_L_ And Res6(Res6.Length - 1) <= B_Light_L_ Then adx3 = 1
                        If (Res6(Res6.Length - 1) + Rang_GSC_) > T_DIM_L_ And (Res6(Res6.Length - 1) + Rang_GSC_) <= B_Light_L_ Or (Res6(Res6.Length - 1) - Rang_GSC_) > T_DIM_L_ And (Res6(Res6.Length - 1) - Rang_GSC_) <= B_Light_L_ Then adx4 = 1

                        If adx1 = 1 And adx2 = 1 And adx4 = 1 And adx3 = 0 Or adx1 = 1 And adx2 = 1 And adx3 = 1 And adx4 = 1 Or adx1 = 0 And adx2 = 1 And adx3 = 1 And adx4 = 1 Then
                            RichText_RES.Text += " - " & __lgry '& vbCrLf
                            cC = cC + 1
                            otp_c = 2
                            '''''''''''''''''''''''''''''''''''''''C-Light_GREY-END''''
                            '''''''''''''''''''''''''''''''''''''''C-Light''''
                        Else
                            adx1 = 0
                            adx2 = 0
                            adx3 = 0
                            adx4 = 0
                            If Res3(Res3.Length - 1) > B_Light_L_ Then adx1 = 1
                            If Res6(Res6.Length - 1) > B_Light_L_ Then adx3 = 1
                            If (Res3(Res3.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx2 = 1
                            If (Res6(Res6.Length - 1) + Rang_GSC_) > B_Light_L_ Then adx4 = 1

                            If adx1 = 1 And adx3 = 1 Or adx1 = 1 And adx3 = 1 And adx2 = 1 And adx4 = 1 Then
                                RichText_RES.Text += " - " & __lght '& vbCrLf
                                vC = vC + 1
                                otp_c = 3

                            Else
                                RichText_RES.Text += " - " & __none '& vbCrLf
                                dC = dC + 1
                                otp_c = 4
                            End If
                        End If
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''C-Light-END''''
                ''''''''''''''''''''''''''''''''''''''' C-END ''''''''
                '
                '''''''''''''''''''''''''''''Gradients''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''GRAD_A'''''''''
                ax = LastNonEmpty

                ''''<lvl>
                If (Res(Res.Length - 1) - Res4(Res4.Length - 1)) >= 0 And (Res(Res.Length - 1) - Res4(Res4.Length - 1)) <= U_N_F_lvl_ Or (Res4(Res4.Length - 1) - Res(Res.Length - 1)) >= 0 And (Res4(Res4.Length - 1) - Res(Res.Length - 1)) <= U_N_F_lvl_ Then
                    RichText_RES2.Text += vbTab & "  " & __lvl & vbTab
                    s = s + 1
                    otp_aG = 9
                Else

                    If Res(Res.Length - 1) < Res4(Res4.Length - 1) Then ax = 1 'DOW Small
                    If (Res(Res.Length - 1) + Rang_GRAD) < Res4(Res4.Length - 1) Then
                        ax = 0 'Dow Big
                    Else
                        If Res4(Res4.Length - 1) < Res(Res.Length - 1) Then ax = 3 ' UP Small
                        If (Res4(Res4.Length - 1) + Rang_GRAD) < Res(Res.Length - 1) Then
                            ax = 2 ' UP Big
                        End If
                    End If

                    If ax = 0 Then 'Dow Big
                        RichText_RES2.Text += vbTab & __Dbig & vbTab
                        m = m + 1
                        otp_aG = 8
                    End If

                    If ax = 1 Then 'DOW Small
                        RichText_RES2.Text += vbTab & __DSm & vbTab
                        n = n + 1
                        otp_aG = 6
                    End If

                    If ax = 2 Then ' UP Big
                        RichText_RES2.Text += vbTab & __TBig & vbTab
                        a = a + 1
                        otp_aG = 7
                    End If
                    If ax = 3 Then ' UP Small
                        RichText_RES2.Text += vbTab & __TSm & vbTab
                        b = b + 1
                        otp_aG = 5
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''GRAD_A_END
                '''''''''''''''''''''''''''''''''''''''GRAD_B''''
                ax = LastNonEmpty
                ''''<lvl>
                If (Res2(Res2.Length - 1) - Res5(Res5.Length - 1)) >= 0 And (Res2(Res2.Length - 1) - Res5(Res5.Length - 1)) <= U_N_F_lvl_ Or (Res5(Res5.Length - 1) - Res2(Res2.Length - 1)) >= 0 And (Res5(Res5.Length - 1) - Res2(Res2.Length - 1)) <= U_N_F_lvl_ Then
                    RichText_RES2.Text += "  " & __lvl & vbTab
                    sB = sB + 1
                    otp_bG = 9
                Else
                    If Res2(Res2.Length - 1) < Res5(Res5.Length - 1) Then ax = 1 'DOW Small
                    If (Res2(Res2.Length - 1) + Rang_GRAD) < Res5(Res5.Length - 1) Then
                        ax = 0 'Dow Big
                    Else
                        If Res5(Res5.Length - 1) < Res2(Res2.Length - 1) Then ax = 3 ' UP Small
                        If (Res5(Res5.Length - 1) + Rang_GRAD) < Res2(Res2.Length - 1) Then
                            ax = 2 ' UP Big
                        End If
                    End If

                    If ax = 0 Then 'Dow Big
                        RichText_RES2.Text += __Dbig & vbTab
                        mB = mB + 1
                        otp_bG = 8
                    End If

                    If ax = 1 Then 'DOW Small
                        RichText_RES2.Text += __DSm & vbTab
                        nB = nB + 1
                        otp_bG = 6
                    End If
                    If ax = 2 Then ' UP Big
                        RichText_RES2.Text += __TBig & vbTab
                        aB = aB + 1
                        otp_bG = 7
                    End If

                    If ax = 3 Then ' UP Small
                        RichText_RES2.Text += __TSm & vbTab
                        bB = bB + 1
                        otp_bG = 5
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''GRAD_B_END
                '''''''''''''''''''''''''''''''''''''''GRAD_C''''
                ax = LastNonEmpty
                ''''<lvl>
                If (Res3(Res3.Length - 1) - Res6(Res6.Length - 1)) >= 0 And (Res3(Res3.Length - 1) - Res6(Res6.Length - 1)) <= U_N_F_lvl_ Or (Res6(Res6.Length - 1) - Res3(Res3.Length - 1)) >= 0 And (Res6(Res6.Length - 1) - Res3(Res3.Length - 1)) <= U_N_F_lvl_ Then
                    RichText_RES2.Text += "  " & __lvl & vbCrLf
                    sC = sC + 1
                    otp_cG = 9
                Else
                    If Res3(Res3.Length - 1) < Res6(Res6.Length - 1) Then ax = 1 'DOW Small
                    If (Res3(Res3.Length - 1) + Rang_GRAD) < Res6(Res6.Length - 1) Then
                        ax = 0 'Dow Big
                    Else
                        If Res6(Res6.Length - 1) < Res3(Res3.Length - 1) Then ax = 3 ' UP Small
                        If (Res6(Res6.Length - 1) + Rang_GRAD) < Res3(Res3.Length - 1) Then
                            ax = 2 ' UP Big
                        End If
                    End If

                    If ax = 0 Then 'Dow Big
                        RichText_RES2.Text += __Dbig & vbCrLf
                        mC = mC + 1
                        otp_cG = 8
                    End If
                    If ax = 1 Then 'DOW Small
                        RichText_RES2.Text += __DSm & vbCrLf
                        nC = nC + 1
                        otp_cG = 6
                    End If
                    If ax = 2 Then ' UP Big
                        RichText_RES2.Text += __TBig & vbCrLf
                        aC = aC + 1
                        otp_cG = 7
                    End If
                    If ax = 3 Then ' UP Small
                        RichText_RES2.Text += __TSm & vbCrLf
                        bC = bC + 1
                        otp_cG = 5
                    End If
                End If
            End If
            '''''''''''''''''''''''''''''GRAD_C_END''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''DEFECT Conditions'''''''''''''''''''''''''
            ' GREY LINE DETECTION CONDITION    ' 5 - SM Top       ' 6 - Sm Bottom       ' 7 - big Top       ' 8 - big Bottom       ' 9 - <->
            If otp_bG = 7 And otp_cG = 8 Or
                otp_aG = 7 And otp_bG = 8 Or
                otp_bG = 8 And otp_cG = 7 Or
                otp_aG = 8 And otp_bG = 7 Or
                otp_aG = 7 And otp_cG = 8 Or
                otp_aG = 8 And otp_cG = 7 Then
                RichText_RES.Text += vbTab & "❷"
                gld = gld + 1
                otp_gli = 13
            Else
                RichText_RES.Text += " "
                otp_gli = 14
            End If
            '''''''''''''''''''''''''''' OUT-OF_POCKET DETECTION CONDITION '''''''''''''''''''''''''''''' 
            If otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 5 And otp_bG = 6 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 4 And otp_aG = 5 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 1 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 4 And otp_aG = 7 And otp_bG = 5 And otp_cG = 7 Or
                    otp_a = 0 And otp_b = 0 And otp_c = 1 And otp_aG = 9 And otp_bG = 6 And otp_cG = 9 Or '5
                    otp_a = 0 And otp_b = 4 And otp_c = 2 And otp_aG = 9 And otp_bG = 8 And otp_cG = 5 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 4 And otp_aG = 7 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 0 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 4 And otp_aG = 5 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 2 And otp_aG = 5 And otp_bG = 7 And otp_cG = 5 Or '10
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 5 And otp_bG = 5 And otp_cG = 9 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 5 And otp_bG = 7 And otp_cG = 9 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 2 And otp_aG = 9 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 7 And otp_bG = 5 And otp_cG = 9 Or
                    otp_a = 2 And otp_b = 1 And otp_c = 0 And otp_aG = 9 And otp_bG = 6 And otp_cG = 9 Or '15
                    otp_a = 2 And otp_b = 1 And otp_c = 0 And otp_aG = 6 And otp_bG = 9 And otp_cG = 9 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 6 And otp_bG = 6 And otp_cG = 5 Or
                    otp_a = 4 And otp_b = 4 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 5 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 1 And otp_aG = 6 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 4 And otp_b = 1 And otp_c = 1 And otp_aG = 7 And otp_bG = 5 And otp_cG = 7 Or '20
                    otp_a = 1 And otp_b = 0 And otp_c = 0 And otp_aG = 5 And otp_bG = 5 And otp_cG = 5 Or
                    otp_a = 2 And otp_b = 4 And otp_c = 0 And otp_aG = 5 And otp_bG = 8 And otp_cG = 9 Or
                    otp_a = 4 And otp_b = 1 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 0 And otp_b = 4 And otp_c = 1 And otp_aG = 6 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 4 And otp_b = 1 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 5 Or '25
                    otp_a = 2 And otp_b = 1 And otp_c = 0 And otp_aG = 5 And otp_bG = 7 And otp_cG = 5 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 9 And otp_bG = 5 And otp_cG = 5 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 9 And otp_bG = 7 And otp_cG = 5 Or
                    otp_a = 2 And otp_b = 1 And otp_c = 0 And otp_aG = 6 And otp_bG = 8 And otp_cG = 9 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 9 And otp_bG = 5 And otp_cG = 7 Or '30
                    otp_a = 0 And otp_b = 1 And otp_c = 2 And otp_aG = 9 And otp_bG = 6 And otp_cG = 9 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 2 And otp_aG = 9 And otp_bG = 9 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 9 And otp_bG = 5 And otp_cG = 5 Or
                    otp_a = 2 And otp_b = 4 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 6 Or '35
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 6 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 2 And otp_b = 4 And otp_c = 1 And otp_aG = 8 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 2 And otp_b = 2 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 6 Or
                    otp_a = 4 And otp_b = 4 And otp_c = 1 And otp_aG = 8 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 4 And otp_b = 4 And otp_c = 1 And otp_aG = 7 And otp_bG = 7 And otp_cG = 7 Or '40
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 5 And otp_bG = 5 And otp_cG = 5 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 5 And otp_bG = 5 And otp_cG = 9 Or
                    otp_a = 1 And otp_b = 2 And otp_c = 2 And otp_aG = 6 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 2 And otp_aG = 6 And otp_bG = 7 And otp_cG = 7 Or '45
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 6 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 2 And otp_aG = 8 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 4 And otp_aG = 8 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 4 And otp_aG = 7 And otp_bG = 7 And otp_cG = 7 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 6 And otp_bG = 8 And otp_cG = 8 Or '50 
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 9 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 4 And otp_b = 1 And otp_c = 0 And otp_aG = 7 And otp_bG = 5 And otp_cG = 5 Or
                    otp_a = 4 And otp_b = 4 And otp_c = 0 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 0 And otp_b = 4 And otp_c = 0 And otp_aG = 6 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 0 And otp_b = 0 And otp_c = 0 And otp_aG = 5 And otp_bG = 5 And otp_cG = 5 Or '55
                    otp_a = 0 And otp_b = 0 And otp_c = 1 And otp_aG = 9 And otp_bG = 9 And otp_cG = 6 Or
                    otp_a = 0 And otp_b = 1 And otp_c = 1 And otp_aG = 9 And otp_bG = 6 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 0 And otp_c = 0 And otp_aG = 6 And otp_bG = 6 And otp_cG = 9 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 5 And otp_bG = 7 And otp_cG = 5 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 7 And otp_bG = 7 And otp_cG = 5 Or '60
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 6 And otp_bG = 6 And otp_cG = 9 Or
                    otp_a = 1 And otp_b = 1 And otp_c = 0 And otp_aG = 9 And otp_bG = 6 And otp_cG = 9 Or
                    otp_a = 1 And otp_b = 4 And otp_c = 1 And otp_aG = 8 And otp_bG = 8 And otp_cG = 8 Or
                    otp_a = 2 And otp_b = 1 And otp_c = 0 And otp_aG = 6 And otp_bG = 6 And otp_cG = 9 Or
                    otp_a = 2 And otp_b = 2 And otp_c = 0 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or '65
                    otp_a = 4 And otp_b = 1 And otp_c = 1 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 4 And otp_b = 4 And otp_c = 2 And otp_aG = 8 And otp_bG = 8 And otp_cG = 6 Or
                    otp_a = 1 And otp_b = 0 And otp_c = 0 And otp_aG = 5 And otp_bG = 6 And otp_cG = 9 Then '68

                RichText_RES.Text += " " & "❶" & vbCrLf
                otp = otp + 1
                otp_otp = 11
            Else
                RichText_RES.Text += " " & vbCrLf
                otp_otp = 12
            End If
            '''''''''''''''''''''''''''''Conditions END''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''Conditions for table''''''''''''''''''''''''''''
            Dim _dark As String = "1 - DARK"
            Dim _dim As String = "2 - DIM"
            Dim _lgry As String = "3 - L-Grey"
            Dim _lght As String = "4 - Light"
            Dim _none As String = "5 - ####"
            Dim _lvl As String = "1 - ↑lvl↓"
            Dim _DSm As String = "2 - DOW Small"
            Dim _TSm As String = "3 - TOP_Small"
            Dim _Dbig As String = "4 - DOW_Big"
            Dim _TBig As String = "5 - TOP_Big"

            If TestArray_CELL.Length > (j) Then
                '''''''#
                table.Rows.Add(num)
                '''''''CELL#
                DataGridView1.Rows.Item(i).Cells(1).Value = TestArray_CELL(i)
                ''''''' horizontal
                If otp_a = 0 Then DataGridView1.Rows.Item(i).Cells(2).Value = _dark
                If otp_a = 1 Then DataGridView1.Rows.Item(i).Cells(2).Value = _dim
                If otp_a = 2 Then DataGridView1.Rows.Item(i).Cells(2).Value = _lgry
                If otp_a = 3 Then DataGridView1.Rows.Item(i).Cells(2).Value = _lght
                If otp_a = 4 Then DataGridView1.Rows.Item(i).Cells(2).Value = _none
                '''''''
                If otp_b = 0 Then DataGridView1.Rows.Item(i).Cells(3).Value = _dark
                If otp_b = 1 Then DataGridView1.Rows.Item(i).Cells(3).Value = _dim
                If otp_b = 2 Then DataGridView1.Rows.Item(i).Cells(3).Value = _lgry
                If otp_b = 3 Then DataGridView1.Rows.Item(i).Cells(3).Value = _lght
                If otp_b = 4 Then DataGridView1.Rows.Item(i).Cells(3).Value = _none
                ''''''''
                If otp_c = 0 Then DataGridView1.Rows.Item(i).Cells(4).Value = _dark
                If otp_c = 1 Then DataGridView1.Rows.Item(i).Cells(4).Value = _dim
                If otp_c = 2 Then DataGridView1.Rows.Item(i).Cells(4).Value = _lgry
                If otp_c = 3 Then DataGridView1.Rows.Item(i).Cells(4).Value = _lght
                If otp_c = 4 Then DataGridView1.Rows.Item(i).Cells(4).Value = _none
                '''''''' vertical
                If otp_aG = 9 Then DataGridView1.Rows.Item(i).Cells(5).Value = _lvl
                If otp_aG = 8 Then DataGridView1.Rows.Item(i).Cells(5).Value = _Dbig
                If otp_aG = 7 Then DataGridView1.Rows.Item(i).Cells(5).Value = _TBig
                If otp_aG = 6 Then DataGridView1.Rows.Item(i).Cells(5).Value = _DSm
                If otp_aG = 5 Then DataGridView1.Rows.Item(i).Cells(5).Value = _TSm
                ''''''''
                If otp_bG = 9 Then DataGridView1.Rows.Item(i).Cells(6).Value = _lvl
                If otp_bG = 8 Then DataGridView1.Rows.Item(i).Cells(6).Value = _Dbig
                If otp_bG = 7 Then DataGridView1.Rows.Item(i).Cells(6).Value = _TBig
                If otp_bG = 6 Then DataGridView1.Rows.Item(i).Cells(6).Value = _DSm
                If otp_bG = 5 Then DataGridView1.Rows.Item(i).Cells(6).Value = _TSm
                '''''''''
                If otp_cG = 9 Then DataGridView1.Rows.Item(i).Cells(7).Value = _lvl
                If otp_cG = 8 Then DataGridView1.Rows.Item(i).Cells(7).Value = _Dbig
                If otp_cG = 7 Then DataGridView1.Rows.Item(i).Cells(7).Value = _TBig
                If otp_cG = 6 Then DataGridView1.Rows.Item(i).Cells(7).Value = _DSm
                If otp_cG = 5 Then DataGridView1.Rows.Item(i).Cells(7).Value = _TSm
                '''''''  otp
                If otp_otp = 11 Then DataGridView1.Rows.Item(i).Cells(8).Value = "OTP - ❶"
                If otp_otp = 12 Then DataGridView1.Rows.Item(i).Cells(8).Value = "-"
                '''''''  GLI
                If otp_gli = 13 Then DataGridView1.Rows.Item(i).Cells(9).Value = "GLI - ❷"
                If otp_gli = 14 Then DataGridView1.Rows.Item(i).Cells(9).Value = "-"
                ''''''' 
                DataGridView1.DataSource = table
            End If
            '''''''''''''''''''''''''''''Conditions END''''''''''''''''''''''''''''
            j = j + 1
            j1 = j1 + 1
            num = num + 1
        Next
        'Remove Last Lines at RICHTEXTBOXes
        Me.RichText_RES.Lines = Me.RichText_RES.Text.Split(New Char() {ControlChars.Lf},
                                                   StringSplitOptions.RemoveEmptyEntries)
        Dim newList2 As List(Of String) = RichText_RES.Lines.ToList
        newList2.RemoveAt(RichText_RES.Lines.Count - 1)
        RichText_RES.Lines = newList2.ToArray

        Me.RichText_RES2.Lines = Me.RichText_RES2.Text.Split(New Char() {ControlChars.Lf},
                                                   StringSplitOptions.RemoveEmptyEntries)
        ''''''''''''''''''''''''''''''''''

        TextAAA.Text = (z.ToString("N0")) & vbTab & (x.ToString("N0")) & vbTab & (c.ToString("N0")) & vbTab & (v.ToString("N0")) & vbTab & (d.ToString("N0")) & vbTab & (b.ToString("N0")) & vbTab & (n.ToString("N0")) & vbTab & (a.ToString("N0")) & vbTab & (m.ToString("N0")) & vbTab & (s.ToString("N0"))
        TextBBB.Text = (zB.ToString("N0")) & vbTab & (xB.ToString("N0")) & vbTab & (cB.ToString("N0")) & vbTab & (vB.ToString("N0")) & vbTab & (dB.ToString("N0")) & vbTab & (bB.ToString("N0")) & vbTab & (nB.ToString("N0")) & vbTab & (aB.ToString("N0")) & vbTab & (mB.ToString("N0")) & vbTab & (sB.ToString("N0"))
        TextCCC.Text = (zC.ToString("N0")) & vbTab & (xC.ToString("N0")) & vbTab & (cC.ToString("N0")) & vbTab & (vC.ToString("N0")) & vbTab & (dC.ToString("N0")) & vbTab & (bC.ToString("N0")) & vbTab & (nC.ToString("N0")) & vbTab & (aC.ToString("N0")) & vbTab & (mC.ToString("N0")) & vbTab & (sC.ToString("N0"))


        z = (z + zB + zC)
        x = (x + xB + xC)
        c = (c + cB + cC)
        v = (v + vB + vC)
        d = (d + dB + dC)
        b = (b + bB + bC)
        n = (n + nB + nC)
        m = (m + mB + mC)
        a = (a + aB + aC)
        s = (s + sB + sC)


        num = num - 2
        z1 = ((z * 100) / (num * 3))  ' Dark
        x1 = ((x * 100) / (num * 3))  ' Dim
        c1 = ((c * 100) / (num * 3))  ' L-Grey
        v1 = ((v * 100) / (num * 3))  ' Light
        d1 = ((d * 100) / (num * 3))  ' Gradi_error

        b1 = ((b * 100) / (num * 3))  ' Sm↑
        n1 = ((n * 100) / (num * 3))  ' Sm↓
        a1 = ((a * 100) / (num * 3))  ' Big↑
        m1 = ((m * 100) / (num * 3))  ' Big↓
        s1 = ((s * 100) / (num * 3))  ' ↑lvl↓


        otp1 = ((otp * 100) / (num))
        gld1 = ((gld * 100) / (num))
        dd = d1 + z1 + x1 + c1 + v1 ' horizontal
        dd2 = b1 + n1 + m1 + a1 + s1 '  vertical

        Rich_Lbl_stat_main.Text = num & vbTab & (z.ToString("N0")) & vbTab & (x.ToString("N0")) & vbTab & (c.ToString("N0")) & vbTab & (v.ToString("N0")) & vbTab & (d.ToString("N0")) & vbTab & (b.ToString("N0")) & vbTab & (n.ToString("N0")) & vbTab & (a.ToString("N0")) & vbTab & (m.ToString("N0")) & vbTab & (s.ToString("N0")) & vbTab & (otp) & vbTab & (gld)

        ReDim Preserve TestArray(LastNonEmpty)
    End Sub
    Private Sub SEPARATOR()

        Dim FirstBox() As String = TextBox_INTENS.Lines
        A_top_text.Text = ""
        B_top_text.Text = ""
        C_top_text.Text = ""
        A_butt_text.Text = ""
        B_butt_text.Text = ""
        C_butt_text.Text = ""
        RichText_RES.Text = ""
        RichText_RES2.Text = ""

        For i = 0 To UBound(FirstBox)
            A_top_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With
            ''''
            i = i + 5
        Next
        For i = 1 To UBound(FirstBox)
            B_top_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With

            i = i + 5
        Next
        For i = 2 To UBound(FirstBox)
            C_top_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With
            i = i + 5
        Next
        For i = 3 To UBound(FirstBox)
            A_butt_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With
            i = i + 5
        Next
        For i = 4 To UBound(FirstBox)
            B_butt_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With
            i = i + 5
        Next
        For i = 5 To UBound(FirstBox)
            C_butt_text.Text &= FirstBox(i) & vbNewLine
            With ProgressBar1 'Progress Bar 
                .Style = ProgressBarStyle.Blocks
                .Step = 0.1
                .Minimum = 0
                .Maximum = i
                .Value = i
            End With
            i = i + 5
        Next
    End Sub


    Private Sub Clean_Click(sender As Object, e As EventArgs)
        TextBox_INTENS.Text = ""
        A_top_text.Text = ""
        B_top_text.Text = ""
        C_top_text.Text = ""
        A_butt_text.Text = ""
        B_butt_text.Text = ""
        C_butt_text.Text = ""
        RichText_RES.Text = ""
        RichText_RES2.Text = ""
        TextBox_CELL_NUM.Text = ""
        TBox3.Text = ""
        txtImageDirectory.Text = ""
        LBL_QTTY.Text = ""
        Rich_Lbl_stat_main.Text = ""
        table.Clear()
        mPreviousImage = Nothing
        mNextImage = Nothing
        mCurrentImage = Nothing
    End Sub
    Private Sub mnuFChildMenus_Click(
ByVal sender As System.Object,
ByVal e As System.EventArgs) _
Handles StatToolStripMenuItem.Click
        Dim xx As Form_stat = New Form_stat
        xx.Show()
        ''''' STAT-Xtic DATA
        xx.Lbl_stat2.Text = " On " & (num * 3) & " SubCells" & vbCrLf & vbCrLf & "horizontal (a+b+c) color" & vbCrLf & " Dark " & "    " & (z.ToString("N0")) & vbCrLf & " Dim " & "    " & (x.ToString("N0")) & vbCrLf & " L-Grey " & "    " & (c.ToString("N0")) & vbCrLf & " Light " & "    " & (v.ToString("N0")) & vbCrLf & "Gradi_error" & "    " & (d.ToString("N0")) & vbCrLf & "SUM" & "    " & (num * 3) & vbCrLf & "______________" & vbCrLf & vbCrLf & "vertical (a+b+c) gradients " & vbCrLf & "Sm↑ " & "    " & (b.ToString("N0")) & vbCrLf & "Sm↓ " & "    " & (n.ToString("N0")) & vbCrLf & "Big↑" & "    " & (a.ToString("N0")) & vbCrLf & "Big↓" & "    " & (m.ToString("N0")) & vbCrLf & "↑lvl↓" & "    " & (s.ToString("N0")) & vbCrLf & vbCrLf & "SUM sub-cell" & "    " & (num * 3) & vbCrLf & "______________" & vbCrLf & " On " & num & " Cells " & vbCrLf & "out-of-pocket" & "    " & (otp) & vbCrLf & "Contact" & "    " & (gld) & vbCrLf
        xx.Lbl_stat.Text = " On " & (num * 3) & " SubCells " & vbCrLf & vbCrLf & "horizontal (a+b+c) color %" & vbCrLf & " Dark % " & "    " & (z1.ToString("N3")) & vbCrLf & " Dim % " & "    " & (x1.ToString("N3")) & vbCrLf & " L-Grey %" & "    " & (c1.ToString("N3")) & vbCrLf & " Light % " & "    " & (v1.ToString("N3")) & vbCrLf & "Gradi_error %" & "    " & (d1.ToString("N3")) & vbCrLf & "SUM horizontal %" & "    " & (dd2.ToString("N3")) & vbCrLf & "______________" & vbCrLf & vbCrLf & "vertical (a+b+c) gradients %  " & vbCrLf & "Sm↑ %" & "    " & (b1.ToString("N3")) & vbCrLf & "Sm↓ %" & "    " & (n1.ToString("N3")) & vbCrLf & "Big↑ %" & "    " & (a1.ToString("N3")) & vbCrLf & "Big↓ %" & "    " & (m1.ToString("N3")) & vbCrLf & "↑lvl↓ %" & "    " & (s1.ToString("N3")) & vbCrLf & vbCrLf & "SUM vertical %" & "    " & (dd2.ToString("N3")) & vbCrLf & "______________" & vbCrLf & " On " & num & " Cells " & vbCrLf & "out-of-pocket %" & "    " & (otp1.ToString("N3")) & vbCrLf & "Contact %" & "    " & (gld1.ToString("N3")) & vbCrLf
        xx.Rich_Lbl_stat.Text = num & vbTab & (z.ToString("N0")) & vbTab & (x.ToString("N0")) & vbTab & (c.ToString("N0")) & vbTab & (v.ToString("N0")) & vbTab & (d.ToString("N0")) & vbTab & (b.ToString("N0")) & vbTab & (n.ToString("N0")) & vbTab & (m.ToString("N0")) & vbTab & (a.ToString("N0")) & vbTab & (s.ToString("N0")) & vbTab & (otp) & vbTab & (gld)
        '''''''CHART
        xx.Chart.Series.Add("Dark")
        xx.Chart.Series.Add("Dim")
        xx.Chart.Series.Add("L-Grey")
        xx.Chart.Series.Add("Light")
        xx.Chart.Series.Add("Gradi_error")
        xx.Chart.Series.Add("Sm↑")
        xx.Chart.Series.Add("Sm↓")
        xx.Chart.Series.Add("Big↑")
        xx.Chart.Series.Add("Big↓")
        xx.Chart.Series.Add("↑lvl↓")
        xx.Chart.Series.Add("out-of-pocket")
        xx.Chart.Series.Add("Contact issue")
        ''''''
        xx.Chart.Series("Dark").Points.Add(z1)
        xx.Chart.Series("Dim").Points.Add(x1)
        xx.Chart.Series("L-Grey").Points.Add(c1)
        xx.Chart.Series("Light").Points.Add(v1)
        xx.Chart.Series("↑lvl↓").Points.Add(s1)
        xx.Chart.Series("Gradi_error").Points.Add(d1)
        xx.Chart.Series("Sm↓").Points.Add(n1)
        xx.Chart.Series("Sm↑").Points.Add(b1)
        xx.Chart.Series("Big↑").Points.Add(a1)
        xx.Chart.Series("Big↓").Points.Add(m1)
        xx.Chart.Series("out-of-pocket").Points.Add(otp1)
        xx.Chart.Series("Contact issue").Points.Add(gld1)
    End Sub
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim FFF As String
        Me.FolderBrowserDialog1.Description = "Select the image directory."  ' Label the folder browser dialog
        Me.FolderBrowserDialog1.ShowNewFolderButton = False   ' disable create folder feature of folder browser
        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()    ' display the folder browser

        If result = DialogResult.OK Then  ' on okay, set the image folder to the
            mFolder = FolderBrowserDialog1.SelectedPath   ' folder browser's selected path
            If (Not String.IsNullOrEmpty(mFolder)) Then
                txtImageDirectory.Text = mFolder
            Else
                Return   ' exit if the user cancels
            End If
            mImageList = New ArrayList()   ' initialize the image arraylist

            Dim Dir As New DirectoryInfo(mFolder)
            'Dim f As FileInfo
            For Each f In Dir.GetFiles("*.PNG*")
                Select Case (f.Extension.ToUpper())
                    'Case ".JPG"
                    '    mImageList.Add(f.FullName)
                    Case ".PNG"
                        mImageList.Add(f.FullName)
                    Case Else
                        ' skip file
                End Select
                FFF = Convert.ToString(f) & vbCrLf
                FFF = Mid(FFF, 24, 8)
                TextBox_CELL_NUM.Text += FFF & vbCrLf
            Next
            mImagePosition = 0
            SetImages()
            '''''last line del
            Dim b As String() = Split(TextBox_CELL_NUM.Text, vbNewLine)
            TextBox_CELL_NUM.Text = String.Join(vbNewLine, b, 0, b.Length - 1)
        End If
    End Sub
    Private Sub SetImages()
        mPreviousImage = Nothing ' clear any existing images
        mNextImage = Nothing ' memory will be an issue if
        mCurrentImage = Nothing ' cycling through a large
        ' set the previous image
        If mImagePosition > 0 Then ' number of images
            Try ' set delegate
                Dim prevCallback As Image.GetThumbnailImageAbort =
                        New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                ' get the previous image
                Dim prevBmp As _
                    New Bitmap(mImageList(mImagePosition - 1).ToString())
                ' thumbnail the image to the size
                ' of the picture box
                mPreviousImage =
                prevBmp.GetThumbnailImage(160, 100,
                        prevCallback, IntPtr.Zero)
                ' set the picture box image
                picPreviousImage.Image = mPreviousImage
                ' clear everything out
                prevBmp = Nothing
                prevCallback = Nothing
                mPreviousImage = Nothing
            Catch
            End Try
        Else
            ' at the limit clear the image
            picPreviousImage.Image = Nothing
        End If
        ' set current image
        If mImagePosition < mImageList.Count Then
            Try
                ' set delegate
                Dim currentCallback =
                        New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                ' get the current image
                Dim currentBmp As _
                        New Bitmap(mImageList(mImagePosition).ToString())
                ' thumbnail the image to the size
                ' of the picture box
                mCurrentImage =
                currentBmp.GetThumbnailImage(450, 450, currentCallback, IntPtr.Zero)
                ' set the picture box image
                PictureBox1.Image = mCurrentImage
                ' clear everything out
                currentBmp = Nothing
                mCurrentImage = Nothing
                currentCallback = Nothing
            Catch
                'stall if it hangs, the user can retry
            End Try
        End If
        ' set next image
        If mImagePosition < mImageList.Count - 1 Then
            Try
                ' set delegate
                Dim nextCallback As _
                    New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                ' get the next image
                Dim nextBmp As _
                    New Bitmap(mImageList(mImagePosition + 1).ToString())
                ' thumbnail the image to the size
                ' of the picture box
                mNextImage =
                nextBmp.GetThumbnailImage(160, 100, nextCallback, IntPtr.Zero)
                ' set the picture box image
                picNextImage.Image = mNextImage
                ' clear everything out
                nextBmp = Nothing
                nextCallback = Nothing
                mNextImage = Nothing
            Catch
                'stall if it hangs, the user can retry
            End Try
        Else
            ' at the limit clear the
            ' image
            picNextImage.Image = Nothing
        End If
        ' call for garbage collection
        GC.Collect()
    End Sub
    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        Dim info_ As Form_info = New Form_info
        info_.Show()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)
        table.DefaultView.Sort = "A_zone-HORIZ, B_zone-HORIZ, C_zone-HORIZ"
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs)
        table.DefaultView.Sort = "A_zone-VERT, B_zone-VERT, C_zone-VERT"
    End Sub
    Private Sub btnOpenImage_Click(sender As Object, e As EventArgs) Handles btnOpenImage.Click

        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else
            System.Diagnostics.Process.Start(mImageList(mImagePosition).ToString())
        End If
    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        '''''''EXIT BUTTON
        Application.Exit()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        ' running ImageJ
        'Dim p As Process() = Process.GetProcessesByName("ImageJ")
        'Process1.StartInfo.FileName = ("C:\Program Files (x86)\ImageJ\ImageJ.exe")
        'Process1.Start()
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        '''''''FORM SIZE CHANGED
        'Dim BIG_Size As Size = New Size(1435, 810)
        'Dim Insize As Size = New Size(1111, 810)
        'If CheckBox1.Checked = True Then
        '    Me.Size = BIG_Size
        'Else Me.Size = Insize
        'End If
    End Sub

    Private Sub SortVerticalGradientToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortVerticalGradientToolStripMenuItem.Click
        table.DefaultView.Sort = "A_zone-VERT, B_zone-VERT, C_zone-VERT"
    End Sub

    Private Sub SortHorizontalGradientToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortHorizontalGradientToolStripMenuItem.Click
        table.DefaultView.Sort = "A_zone-HORIZ, B_zone-HORIZ, C_zone-HORIZ"
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        TextBox_INTENS.Text = ""
        A_top_text.Text = ""
        B_top_text.Text = ""
        C_top_text.Text = ""
        A_butt_text.Text = ""
        B_butt_text.Text = ""
        C_butt_text.Text = ""
        RichText_RES.Text = ""
        RichText_RES2.Text = ""
        TextBox_CELL_NUM.Text = ""
        TBox3.Text = ""
        txtImageDirectory.Text = ""
        LBL_QTTY.Text = ""
        Rich_Lbl_stat_main.Text = ""
        table.Clear()
        mPreviousImage = Nothing
        mNextImage = Nothing
        mCurrentImage = Nothing

        PictureBox1.Image = Nothing
        PictureBox1.BackColor = Color.Empty
        PictureBox1.Invalidate()

    End Sub

    Private Sub RUNALLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RUNALLToolStripMenuItem.Click
        '''''''''''
        '''''''''''
        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else

            Dim qutty As Integer = (mImageList.Count)
            For xxx As Integer = 1 To qutty
                Randomize()
                Dim sum As Int32
                Dim sum2 As Int32
                Dim sum3 As Int32
                Dim sum4 As Int32
                Dim sum5 As Int32
                Dim sum6 As Int32
                Dim rnd As New Random()
                Dim b As Bitmap = PictureBox1.Image
                Dim j As Integer = 300
                For i As Integer = 1 To j Step 1
                    '''''TOP A
                    Dim pixelcolor As Color
                    Dim value As Integer
                    Dim value2 As Integer
                    Dim index As Integer = 0
                    Do
                        value = CInt(Int((rnd.Next(40, 115)) + 1))
                        value2 = CInt(Int((rnd.Next(30, 139)) + 1))
                        pixelcolor = b.GetPixel(value2, value) 'choose pix color 
                    Loop Until pixelcolor <> Color.DarkRed
                    b.SetPixel(value2, value, Color.DarkRed) 'marked choosen pix in red
                    '''''TOP B
                    Dim pixelcolor2 As Color
                    Dim value3 As Integer
                    Dim value4 As Integer
                    Do
                        value3 = CInt(Int((rnd.Next(40, 115)) + 1))
                        value4 = CInt(Int((rnd.Next(175, 290)) + 1))
                        pixelcolor2 = b.GetPixel(value4, value3) 'choose pix color 
                    Loop Until pixelcolor2 <> Color.DarkRed
                    b.SetPixel(value4, value3, Color.DarkRed) 'marked choosen pix in red
                    ''''' TOP C
                    Dim pixelcolor3 As Color
                    Dim value5 As Integer
                    Dim value6 As Integer
                    Do
                        value5 = CInt(Int((rnd.Next(40, 115)) + 1))
                        value6 = CInt(Int((rnd.Next(325, 432)) + 1))
                        pixelcolor3 = b.GetPixel(value6, value5) 'choose pix color 
                    Loop Until pixelcolor3 <> Color.DarkRed
                    b.SetPixel(value6, value5, Color.DarkRed) 'marked choosen pix in red
                    ''''''BOTTOM A
                    Dim pixelcolor4 As Color
                    Dim value7 As Integer
                    Dim value8 As Integer
                    Do
                        value7 = CInt(Int((rnd.Next(338, 415)) + 1))
                        value8 = CInt(Int((rnd.Next(30, 139)) + 1))
                        pixelcolor4 = b.GetPixel(value8, value7) 'choose pix color 
                    Loop Until pixelcolor4 <> Color.DarkRed
                    b.SetPixel(value8, value7, Color.DarkRed) 'marked choosen pix in red
                    ''''' BOTTOM B
                    Dim pixelcolor5 As Color
                    Dim value9 As Integer
                    Dim value10 As Integer
                    Do
                        value9 = CInt(Int((rnd.Next(338, 415)) + 1))
                        value10 = CInt(Int((rnd.Next(175, 290)) + 1))
                        pixelcolor5 = b.GetPixel(value10, value9) 'choose pix color 
                    Loop Until pixelcolor5 <> Color.DarkRed
                    b.SetPixel(value10, value9, Color.DarkRed) 'marked choosen pix in red
                    ''''' BOTTOM C
                    Dim pixelcolor6 As Color
                    Dim value11 As Integer
                    Dim value12 As Integer
                    Do
                        value11 = CInt(Int((rnd.Next(338, 415)) + 1))
                        value12 = CInt(Int((rnd.Next(325, 432)) + 1)) 'x
                        pixelcolor6 = b.GetPixel(value12, value11) 'choose pix color 
                    Loop Until pixelcolor6 <> Color.DarkRed
                    b.SetPixel(value12, value11, Color.DarkRed) 'marked choosen pix in red

                    sum = sum + pixelcolor.B
                    sum2 = sum2 + pixelcolor2.B
                    sum3 = sum3 + pixelcolor3.B
                    sum4 = sum4 + pixelcolor4.B
                    sum5 = sum5 + pixelcolor5.B
                    sum6 = sum6 + pixelcolor6.B
                Next i
                sum = sum / j
                sum2 = sum2 / j
                sum3 = sum3 / j
                sum4 = sum4 / j
                sum5 = sum5 / j
                sum6 = sum6 / j
                TextBox_INTENS.Text += Convert.ToString(sum) & vbCrLf & Convert.ToString(sum2) & vbCrLf & Convert.ToString(sum3) & vbCrLf & Convert.ToString(sum4) & vbCrLf & Convert.ToString(sum5) & vbCrLf & Convert.ToString(sum6) & vbCrLf
                ''''''' RED Rectangles
                Dim p As New System.Drawing.Pen(Color.Red, 2)
                Dim g As System.Drawing.Graphics
                PictureBox1.Refresh()
                g = PictureBox1.CreateGraphics
                '''''TOP
                g.DrawRectangle(p, 30, 40, 109, 75)
                g.DrawRectangle(p, 175, 40, 115, 75)
                g.DrawRectangle(p, 325, 40, 109, 75)
                '''''BOTTOM
                g.DrawRectangle(p, 30, 340, 109, 75)
                g.DrawRectangle(p, 175, 340, 115, 75)
                g.DrawRectangle(p, 325, 340, 109, 75)

                If (mImagePosition <= (mImageList.Count - 2)) Then
                    mImagePosition += 1
                    SetImages()
                End If
                LBL_QTTY.Text = xxx & " Cells"
            Next
            '''''last line del
            Dim bbb As String() = Split(TextBox_INTENS.Text, vbNewLine)
            TextBox_INTENS.Text = String.Join(vbNewLine, bbb, 0, bbb.Length - 1)


            '''''' SEPARATION
            SEPARATOR()
            '''''' Calculation
            MAIN_CALC()
        End If
    End Sub
    Private Sub SeparationToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SeparationToolStripMenuItem1.Click
        SEPARATOR()
    End Sub

    Private Sub RunCalcToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RunCalcToolStripMenuItem.Click

        ''''''Protection from idionts 

        Dim lineCount As Integer = TextBox_INTENS.Lines.Length - 1
        Dim lineCount2 As Integer = TextBox_CELL_NUM.Lines.Length

        If lineCount2 <> (lineCount / 6) Then
            MsgBox("Cell field empty! Input cell number" & vbCrLf & "You make" & " " & (lineCount / 6) & " measurements" & "But insert only " & "  " & lineCount2 & " cells numbers")
        Else
            MAIN_CALC()
        End If


    End Sub

    Private Sub ColorDetectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorDetectionToolStripMenuItem.Click
        ' Get a Bitmap from the PictureBox
        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else
            TBox3.Text = ""
            Dim sum As Integer
            Dim sum2 As Integer
            Dim sum3 As Integer
            Dim sum4 As Integer
            Dim Resolut As Integer = 202500
            Dim bm As Bitmap = PictureBox1.Image
            Dim colorList As New List(Of System.Drawing.Color)
            Dim groups = colorList.GroupBy(Function(value) value).OrderByDescending(Function(g) g.Count)
            For x As Integer = 0 To bm.Width - 1
                For y As Integer = 0 To bm.Height - 1
                    colorList.Add(bm.GetPixel(x, y))
                Next
            Next
            For Each grp In groups
                If Convert.ToInt32(grp(0).R) >= 0 And Convert.ToInt32(grp(0).R) <= 60 Then
                    sum += grp.Count
                End If
                If Convert.ToInt32(grp(0).R) >= 61 And Convert.ToInt32(grp(0).R) <= 140 Then
                    sum2 += grp.Count
                End If
                If Convert.ToInt32(grp(0).R) >= 141 And Convert.ToInt32(grp(0).R) <= 200 Then
                    sum3 += grp.Count
                End If
                If Convert.ToInt32(grp(0).R) >= 201 And Convert.ToInt32(grp(0).R) <= 255 Then
                    sum4 += grp.Count
                End If
                TBox3.Text = (" 0-60 " & "        - " & ((sum * 100) / Resolut).ToString("N2")) & "%" & vbTab & (" 61-140 " & "   - " & ((sum2 * 100) / Resolut).ToString("N2")) & "%" & vbTab & (" 141-200 " & " - " & ((sum3 * 100) / Resolut).ToString("N2")) & "%" & vbTab & (" 201-255 " & " - " & ((sum4 * 100) / Resolut).ToString("N2")) & "%"
            Next
            'cell_name.Text = ""
            ' cell_name.Text = f.ToString(mImagePosition.ToString()) '(mImagePosition)

            ' cell_name.Text = Mid((mImageList(mImagePosition).ToString()), 14, 35)
        End If
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else
            If (mImagePosition >= 0 And Not picPreviousImage.Image Is Nothing) Then
                mImagePosition -= 1
                SetImages()
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else
            If (mImagePosition <= (mImageList.Count - 2)) Then
                mImagePosition += 1
                SetImages()
            End If
        End If
    End Sub

    Private Sub ManualRunImageProssToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManualRunImageProssToolStripMenuItem.Click
        If TextBox_CELL_NUM.Text = "" Then
            MsgBox("Cell fields empty! Choose folder with EL's first ")
        Else

            Randomize()
            Dim sum As Int32
            Dim sum2 As Int32
            Dim sum3 As Int32
            Dim sum4 As Int32
            Dim sum5 As Int32
            Dim sum6 As Int32
            Dim rnd As New Random()
            Dim b As Bitmap = PictureBox1.Image
            Dim j As Integer = 300
            For i As Integer = 1 To j Step 1
                '''''TOP A
                Dim pixelcolor As Color
                Dim value As Integer
                Dim value2 As Integer
                Dim index As Integer = 0
                Do
                    value = CInt(Int((rnd.Next(40, 115)) + 1))
                    value2 = CInt(Int((rnd.Next(30, 139)) + 1))
                    pixelcolor = b.GetPixel(value2, value) 'choose pix color 
                Loop Until pixelcolor <> Color.DarkRed
                b.SetPixel(value2, value, Color.DarkRed) 'marked choosen pix in red
                '''''TOP B
                Dim pixelcolor2 As Color
                Dim value3 As Integer
                Dim value4 As Integer
                Do
                    value3 = CInt(Int((rnd.Next(40, 115)) + 1))
                    value4 = CInt(Int((rnd.Next(175, 290)) + 1))
                    pixelcolor2 = b.GetPixel(value4, value3) 'choose pix color 
                Loop Until pixelcolor2 <> Color.DarkRed
                b.SetPixel(value4, value3, Color.DarkRed) 'marked choosen pix in red
                ''''' TOP C
                Dim pixelcolor3 As Color
                Dim value5 As Integer
                Dim value6 As Integer
                Do
                    value5 = CInt(Int((rnd.Next(40, 115)) + 1))
                    value6 = CInt(Int((rnd.Next(325, 432)) + 1))
                    pixelcolor3 = b.GetPixel(value6, value5) 'choose pix color 
                Loop Until pixelcolor3 <> Color.DarkRed
                b.SetPixel(value6, value5, Color.DarkRed) 'marked choosen pix in red
                ''''''BOTTOM A
                Dim pixelcolor4 As Color
                Dim value7 As Integer
                Dim value8 As Integer
                Do
                    value7 = CInt(Int((rnd.Next(338, 415)) + 1))
                    value8 = CInt(Int((rnd.Next(30, 139)) + 1))
                    pixelcolor4 = b.GetPixel(value8, value7) 'choose pix color 
                Loop Until pixelcolor4 <> Color.DarkRed
                b.SetPixel(value8, value7, Color.DarkRed) 'marked choosen pix in red
                ''''' BOTTOM B
                Dim pixelcolor5 As Color
                Dim value9 As Integer
                Dim value10 As Integer
                Do
                    value9 = CInt(Int((rnd.Next(338, 415)) + 1))
                    value10 = CInt(Int((rnd.Next(175, 290)) + 1))
                    pixelcolor5 = b.GetPixel(value10, value9) 'choose pix color 
                Loop Until pixelcolor5 <> Color.DarkRed
                b.SetPixel(value10, value9, Color.DarkRed) 'marked choosen pix in red
                ''''' BOTTOM C
                Dim pixelcolor6 As Color
                Dim value11 As Integer
                Dim value12 As Integer
                Do
                    value11 = CInt(Int((rnd.Next(338, 415)) + 1))
                    value12 = CInt(Int((rnd.Next(325, 432)) + 1)) 'x
                    pixelcolor6 = b.GetPixel(value12, value11) 'choose pix color 
                Loop Until pixelcolor6 <> Color.DarkRed
                b.SetPixel(value12, value11, Color.DarkRed) 'marked choosen pix in red

                sum = sum + pixelcolor.B
                sum2 = sum2 + pixelcolor2.B
                sum3 = sum3 + pixelcolor3.B
                sum4 = sum4 + pixelcolor4.B
                sum5 = sum5 + pixelcolor5.B
                sum6 = sum6 + pixelcolor6.B
            Next i
            sum = sum / j
            sum2 = sum2 / j
            sum3 = sum3 / j
            sum4 = sum4 / j
            sum5 = sum5 / j
            sum6 = sum6 / j
            TextBox_INTENS.Text += Convert.ToString(sum) & vbCrLf & Convert.ToString(sum2) & vbCrLf & Convert.ToString(sum3) & vbCrLf & Convert.ToString(sum4) & vbCrLf & Convert.ToString(sum5) & vbCrLf & Convert.ToString(sum6) & vbCrLf
            ''''''' RED Rectangles
            Dim p As New System.Drawing.Pen(Color.Red, 2)
            Dim g As System.Drawing.Graphics
            PictureBox1.Refresh()
            g = PictureBox1.CreateGraphics
            '''''TOP
            g.DrawRectangle(p, 30, 40, 109, 75)
            g.DrawRectangle(p, 175, 40, 115, 75)
            g.DrawRectangle(p, 325, 40, 109, 75)
            '''''BOTTOM
            g.DrawRectangle(p, 30, 340, 109, 75)
            g.DrawRectangle(p, 175, 340, 115, 75)
            g.DrawRectangle(p, 325, 340, 109, 75)
        End If
    End Sub
    Private Sub SmallToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SmallToolStripMenuItem.Click
        ''''''FORM SIZE CHANGED
        Dim Insize As Size = New Size(1368, 559)
        Me.Size = Insize
    End Sub

    Private Sub BigToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BigToolStripMenuItem.Click
        ''''''FORM SIZE CHANGED
        Dim BIG_Size As Size = New Size(1603, 559)
        Me.Size = BIG_Size
    End Sub
End Class