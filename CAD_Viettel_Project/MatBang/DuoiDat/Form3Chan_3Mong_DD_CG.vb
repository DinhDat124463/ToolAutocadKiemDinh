
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry 'Namespace chứa các đối tượng tạo nên thực thể đồ họa
Imports System.IO


' Add ribbon 
Imports Autodesk.Windows
Imports System.Windows.Input
Public Class Form3Chan_3Mong_DD_CG
    Private Sub btnlai_Click(sender As Object, e As EventArgs) Handles btnlai.Click
        '3 chan 3  mong loai 2
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim curUCSMatrix As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem
        Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim TextHight As Double = TiLeChu
        Dim Linetylescale As Double = Tile
        Dim Dimscale As Double = TextHight

        mbVeMong.Ve_Mong_0(b_b0mong, b_h0mong, x1, y1, "1", 10, TextHight)
        mbVeMong.Ve_Cot_Tam_Giac_update(b_a, x1, y1, "5", "2", 10, Dimscale, TextHight, x1, y1, x2, y2, x3, y3, x4, y4, 3)
        mbVeMong.Ga_Chong_Xoay_MB_TamGiac(b_a, x1, y1, "0", "2", 10, Dimscale, TextHight, x1, y1, x2, y2, x3, y3, x4, y4, 3)
        If ThongTinChung.SoMong = 3 Then
            mbVeMong.VeMong(x1, y1, b_bMong1, b_hMong1, b_a, b_bmove, b_hmove, "Móng M1", TextHight, 10, "1", "Dưới đất", b_Mong1)
            mbVeMong.VeMong(x2, y2, b_bMong2, b_hMong2, b_a, b_bmove, b_hmove, "Móng M2", TextHight, 10, "1", "Dưới đất", b_Mong2)
            mbVeMong.VeMong(x3, y3, b_bMong3, b_hMong3, b_a, b_bmove, b_hmove, "Móng M3", TextHight, 10, "1", "Dưới đất", b_Mong3)
        End If
        If ThongTinChung.SoMong = 6 Then
            mbVeMong.Ga_Chong_Xoay_MB_TamGiac(b_a, x1_phu, y1_phu, "0", "2", 10, Dimscale, TextHight, x1_phu, y1_phu, x2_phu, y2_phu, x3_phu, y3_phu, x4_phu, y4_phu, 3)
            mbVeMong.Ve_Cot_Tam_Giac_update_Mongphu(b_a, x1_phu, y1_phu, "5", "2", 10, Dimscale, TextHight, x1_phu, y1_phu, x2_phu, y2_phu, x3_phu, y3_phu, 3)
            mbVeMong.VeMong(x1, y1, b_bMong1, b_hMong1, b_a, b_bmove, b_hmove, "Móng M1", TextHight, 10, "1", "Dưới đất", b_Mong1)
            mbVeMong.VeMong(x2, y2, b_bMong2, b_hMong2, b_a, b_bmove, b_hmove, "Móng M2", TextHight, 10, "1", "Dưới đất", b_Mong2)
            mbVeMong.VeMong(x3, y3, b_bMong3, b_hMong3, b_a, b_bmove, b_hmove, "Móng M3", TextHight, 10, "1", "Dưới đất", b_Mong3)
            mbVeMong.VeMong(x1_phu, y1_phu, b_bmongphu1, b_hmongphu1, b_a, b_bmove, b_hmove, "Móng M4", TextHight, 10, "1", "Dưới đất", b_Mongphu1)
            mbVeMong.VeMong(x2_phu, y2_phu, b_bmongphu2, b_hmongphu2, b_a, b_bmove, b_hmove, "Móng M5", TextHight, 10, "1", "Dưới đất", b_Mongphu2)
            mbVeMong.VeMong(x3_phu, y3_phu, b_bmongphu3, b_hmongphu3, b_a, b_bmove, b_hmove, "Móng M6", TextHight, 10, "1", "Dưới đất", b_Mongphu3)
        End If
        MsgBox("Đã vẽ xong! Tắt bỏ để xem")
    End Sub
End Class