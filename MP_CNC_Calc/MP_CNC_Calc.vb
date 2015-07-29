Imports System.Drawing.Drawing2D

Public Class MP_CNC_Calc
    Dim X_Conduit_Len As Single
    Dim Y_Conduit_Len As Single
    Dim Z_Conduit_Len As Single
    Private Sub btn_Calculate_Click_1(sender As Object, e As EventArgs) Handles btn_Calculate.Click
        Dim X_Len As Single
        Dim Y_Len As Single
        Dim Z_Len As Single

        Dim X_Conduit_Len_x3 As Single
        Dim Y_Conduit_Len_x3 As Single
        Dim Z_Conduit_Len_x2 As Single
        Dim XYZ_Conduit_Len As Single
        Dim X_Belt_Len As Single
        Dim Y_Belt_Len As Single
        Dim X_Belt_Len_x2 As Single
        Dim Y_Belt_Len_x2 As Single
        Dim XY_Belt_Len As Single
        Dim S_XY_Belt_LenFt As String = ""
        Dim Z_Rod_Len As Single
        X_Len = Convert.ToSingle(txt_X.Text)
        Y_Len = Convert.ToSingle(txt_Y.Text)
        Z_Len = Convert.ToSingle(txt_Z.Text)
        X_Conduit_Len = CSng(X_Len + Convert.ToSingle(txt_XCadd.Text))      'Default is 11.0
        Y_Conduit_Len = CSng(Y_Len + Convert.ToSingle(txt_YCadd.Text))      'Default is 11.0
        Z_Conduit_Len = CSng(Z_Len + Convert.ToSingle(txt_ZCadd.Text))      'Default is 7.9
        X_Conduit_Len_x3 = X_Conduit_Len * 3
        Y_Conduit_Len_x3 = Y_Conduit_Len * 3
        Z_Conduit_Len_x2 = Z_Conduit_Len * 2
        XYZ_Conduit_Len = X_Conduit_Len_x3 + Y_Conduit_Len_x3 + Z_Conduit_Len_x2
        X_Belt_Len = X_Conduit_Len + Convert.ToSingle(txt_XBadd.Text)      'Default is 2
        Y_Belt_Len = Y_Conduit_Len + Convert.ToSingle(txt_YBadd.Text)      'Default is 2
        X_Belt_Len_x2 = X_Belt_Len * 2
        Y_Belt_Len_x2 = Y_Belt_Len * 2
        XY_Belt_Len = X_Belt_Len_x2 + Y_Belt_Len_x2
        Z_Rod_Len = Z_Conduit_Len - 2

        txt_X_Conduits.Text = Convert.ToString(X_Conduit_Len & """")           'Lengths of each X-axis pipe
        txt_X_TotC_Len.Text = Convert.ToString(X_Conduit_Len_x3 & """")        'Total Length of X-axis Pipes

        txt_Y_Conduits.Text = Convert.ToString(Y_Conduit_Len & """")           'Lengths of each Y-axis pipe
        txt_Y_TotC_Len.Text = Convert.ToString(Y_Conduit_Len_x3 & """")        'Total Length of Y-axis Pipes

        txt_Z_Conduits.Text = Convert.ToString(Z_Conduit_Len & """")           'Lengths of each Z-axis pipe
        txt_Z_TotC_Len.Text = Convert.ToString(Z_Conduit_Len_x2 & """")        'Total Length of Z-axis Pipes

        txt_Tot_Conduit_InLen.Text = String.Format("{0:f2}", XYZ_Conduit_Len & """")           'Total Length of all pipes in Inches
        'txt_Tot_Conduit_FtLen.Text = String.Format("{0:f4}", XYZ_Conduit_Len / 12.0)    'Total Length of all pipes in Feet
        txt_Tot_Conduit_FtLen.Text = ToFeetAndInches(CSng(CDbl(XYZ_Conduit_Len)))

        txt_X_BeltLen.Text = Convert.ToString(X_Belt_Len & """")               'Length of X-axis Belts
        txt_X_TotB_Len.Text = Convert.ToString(X_Belt_Len_x2 & """")
        txt_Y_BeltLen.Text = Convert.ToString(Y_Belt_Len & """")               'Length of Y-axis Belts
        txt_Y_TotB_Len.Text = Convert.ToString(Y_Belt_Len_x2 & """")
        txt_Z_RodLen.Text = Convert.ToString(Z_Rod_Len & """")                 'Length of Z-axis Threaded Rod

        txt_Tot_Belt_FtLen.Text = ToFeetAndInches(CSng(CDbl(XY_Belt_Len)))
        DrawLayout()
    End Sub
    Public Shared Function FeetInches(ByVal TotalInches As Integer) As String
        Return String.Format("{0} ft. {1} inches", TotalInches \ 12, TotalInches Mod 12)
    End Function
    Private Function ToFeetAndInches(ByVal aDouble As Single) As String
        Dim result As String = ""
        Dim feet, inches As Single
        feet = CLng(aDouble) \ 12
        inches = aDouble Mod 12
        result = feet.ToString & "' - " & String.Format("{0:f4}", inches) & """"
        Return result
    End Function
    Private Sub btn_Defaults_Click_1(sender As Object, e As EventArgs) Handles btn_Defaults.Click
        txt_X.Text = "21"
        txt_Y.Text = "21"
        txt_Z.Text = "6.1"
        txt_XCadd.Text = "11"
        txt_YCadd.Text = "11"
        txt_ZCadd.Text = "11"
        txt_XBadd.Text = "2"
        txt_YBadd.Text = "2"
    End Sub
    Private Sub opt_Disable_CheckedChanged_1(sender As Object, e As EventArgs) Handles opt_Disable.CheckedChanged
        grp_Constants.Enabled = False
    End Sub
    Private Sub opt_Enable_CheckedChanged_1(sender As Object, e As EventArgs) Handles opt_Enable.CheckedChanged
        grp_Constants.Enabled = True
    End Sub

    Private Sub txt_X_TextChanged(sender As Object, e As EventArgs) Handles txt_X.TextChanged
        txt_X1.Text = txt_X.Text
    End Sub

    Private Sub txt_X1_TextChanged(sender As Object, e As EventArgs) Handles txt_X1.TextChanged
        txt_X.Text = txt_X1.Text
    End Sub

    Private Sub txt_Y_TextChanged(sender As Object, e As EventArgs) Handles txt_Y.TextChanged
        txt_Y1.Text = txt_Y.Text
    End Sub

    Private Sub txt_Y1_TextChanged(sender As Object, e As EventArgs) Handles txt_Y1.TextChanged
        txt_Y.Text = txt_Y1.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DrawLayout()
        
    End Sub
    Private Sub DrawLayout()
        'Multiply all dimensions by 10
        Dim Len As Integer
        Dim Wid As Integer
        Dim Center_X As Integer
        Dim Center_Y As Integer
        Dim CNC_X As Integer
        Dim CNC_Y As Integer
        Len = CInt(txt_X.Text) * 10
        Wid = CInt(txt_Y.Text) * 10
        CNC_X = Len + 110
        CNC_Y = Wid + 110
        Center_X = CInt((CNC_X / 2) + 10)
        Center_Y = CInt((CNC_Y / 2) + 10)

        Dim EMT_OD As Integer = 10
        Dim CornerBlk As Integer = 21       'Corner Blocks are 2" square
        Dim RollerBlk As Integer = 31       'X & Y edge roller blocks are 3" square
        Dim MidBlk As Integer = 50          'Middle Block is 5" square
        Dim CNC_Rec As New Rectangle(0, 0, CNC_X, CNC_Y)
        Dim CorBlkRec1 As New Rectangle(10, 10, CornerBlk, CornerBlk)
        Dim CorBlkRec2 As New Rectangle(10, (CNC_Y - 20) + 10, CornerBlk, CornerBlk)
        Dim CorBlkRec3 As New Rectangle((CNC_X - 20) + 10, 10, CornerBlk, CornerBlk)
        Dim CorBlkRec4 As New Rectangle((CNC_X - 20) + 10, (CNC_Y - 20) + 10, CornerBlk, CornerBlk)

        Dim RolBlkRec1 As New Rectangle(5, CInt((CNC_Y / 2) - 5), RollerBlk, RollerBlk)
        Dim RolBlkRec2 As New Rectangle(CNC_X - 15, CInt((CNC_Y / 2) - 5), RollerBlk, RollerBlk)
        Dim RolBlkRec3 As New Rectangle(CInt((CNC_X / 2) - 5), 5, RollerBlk, RollerBlk)
        Dim RolBlkRec4 As New Rectangle(CInt((CNC_X / 2) - 5), CNC_Y - 15, RollerBlk, RollerBlk)
        Dim MidBlkRec As New Rectangle(CInt(Center_X - (MidBlk / 2)), CInt(Center_Y - (MidBlk / 2)), MidBlk, MidBlk)

        Dim Workspace As New Rectangle(CInt(((CNC_X - Len) / 2) + 10), CInt(((CNC_Y - Wid) / 2) + 10), Len, Wid)
        Dim Pipe_X_Top As New Rectangle(10, 15, CNC_X, EMT_OD)
        Dim Pipe_X_Mid As New Rectangle(10, CInt((CNC_Y / 2) + 5), CNC_X, EMT_OD)
        Dim Pipe_X_Bot As New Rectangle(10, (CNC_Y - 15) + 10, CNC_X, EMT_OD)
        Dim Pipe_Y_Lt As New Rectangle(15, 10, EMT_OD, CNC_Y)
        Dim Pipe_Y_Mid As New Rectangle(CInt((CNC_X / 2) + 5), 10, EMT_OD, CNC_Y)
        Dim Pipe_Y_Rt As New Rectangle((CNC_X - 15) + 10, 10, EMT_OD, CNC_Y)

        Dim picWidth As Integer
        Dim PicHeight As Integer
        picWidth = pic_MP_CNC.Width
        PicHeight = pic_MP_CNC.Height
        Dim B As Bitmap = New Bitmap(picWidth, PicHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim G As Graphics = Graphics.FromImage(B)
        Dim fillRect As New Rectangle(0, 0, picWidth, PicHeight)
        G.FillRectangle(Brushes.White, fillRect)
        Dim redPen As Pen = New Pen(Color.Red, 2)
        Dim GreenPen As Pen = New Pen(Color.Green, 2)
        Dim br0 As New HatchBrush(HatchStyle.Max, Color.Black, Color.Black)
        Dim br1 As New HatchBrush(HatchStyle.Max, Color.Red, Color.Red)
        Dim Pt1 As Point = New Point(0, 0)
        Dim Pt2 As Point = New Point(Center_X, Center_Y)
        Dim ClearPic As New Rectangle(0, 0, pic_MP_CNC.Width, pic_MP_CNC.Height)
        Debug.WriteLine((CNC_Y / 2) + 5)
        G.FillRectangle(Brushes.White, ClearPic)
        G.FillRectangle(Brushes.LightGray, Workspace)
        G.FillRectangle(br1, CorBlkRec1)        'Draw Top Left Corner
        G.FillRectangle(br1, CorBlkRec2)        'Draw Bottom Left Corner
        G.FillRectangle(br1, CorBlkRec3)        'Draw Top Right Corner
        G.FillRectangle(br1, CorBlkRec4)        'Draw Bottom Right Corner
        G.FillRectangle(br1, RolBlkRec1)        'Draw Middle Left Roller Block
        G.FillRectangle(br1, RolBlkRec2)        'Draw Middle Right Roller Block
        G.FillRectangle(br1, RolBlkRec3)        'Draw Middle Top Roller Block
        G.FillRectangle(br1, RolBlkRec4)        'Draw Middle Bottom Roller Block
        G.FillRectangle(br1, MidBlkRec)         'Draw Middle Z Block
        G.FillRectangle(br0, Pipe_X_Top)        'Draw Top X Pipe
        G.FillRectangle(br0, Pipe_X_Mid)        'Draw Middle X Pipe
        G.FillRectangle(br0, Pipe_X_Bot)        'Draw Bottom X Pipe

        G.FillRectangle(br0, Pipe_Y_Lt)         'Draw Left Y Pipe
        G.FillRectangle(br0, Pipe_Y_Mid)        'Draw Middle Y Pipe
        G.FillRectangle(br0, Pipe_Y_Rt)         'Draw Right Y Pipe
        pic_MP_CNC.Image = B                    'Draw Graphics to Window
        G.Dispose()
    End Sub

    Private Sub AboutMPCNCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutMPCNCToolStripMenuItem.Click
        About_MP_CNC_Calc.Show()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub opt_Belt2_CheckedChanged(sender As Object, e As EventArgs) Handles opt_Belt2.CheckedChanged
        txt_XBadd.Text = "2"
        txt_YBadd.Text = "2"
    End Sub

    Private Sub opt_Belt5_CheckedChanged(sender As Object, e As EventArgs) Handles opt_Belt5.CheckedChanged
        txt_XBadd.Text = "5"
        txt_YBadd.Text = "5"
    End Sub

    Private Sub opt_Belt7_CheckedChanged(sender As Object, e As EventArgs) Handles opt_Belt7.CheckedChanged
        txt_XBadd.Text = "7"
        txt_YBadd.Text = "7"
    End Sub
End Class
