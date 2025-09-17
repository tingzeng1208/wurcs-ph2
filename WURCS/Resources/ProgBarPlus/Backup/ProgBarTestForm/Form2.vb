Public Class Form2


    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        ProgBarPlus1.Value = ProgBarPlus1.Max - TrackBar1.Value
        ProgBarPlus2.Value = TrackBar1.Value
        ProgBarPlus3.Value = ProgBarPlus1.Max - TrackBar1.Value
        ProgBarPlus4.Value = TrackBar1.Value
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class