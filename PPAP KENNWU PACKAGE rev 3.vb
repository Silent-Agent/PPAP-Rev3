
Private Sub AAR_box_Click()

If AAR_box.Value = True Then
AAR.Visible = xlSheetVisible
Range("J33").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J33").Interior.Color = RGB(146, 208, 80) 'Set the fill color
Else
AAR.Visible = xlSheetHidden
Range("J33").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J33").Interior.Color = RGB(255, 0, 0) 'Set the fill color
End If

End Sub

Private Sub CEA_box_Click()

If CEA_box.Value = True Then
CEA.Visible = xlSheetVisible
Range("J11").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J11").Interior.Color = RGB(146, 208, 80)
Else
CEA.Visible = xlSheetHidden
Range("J11").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J11").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub CP_box_Click()

If CP_box.Value = True Then
CP.Visible = xlSheetVisible
Range("J29").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J29").Interior.Color = RGB(146, 208, 80)
Else
CP.Visible = xlSheetHidden
Range("J29").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J29").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub CSR_box_Click()

If CSR_box.Value = True Then
CSR.Visible = xlSheetVisible
Range("J41").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J41").Interior.Color = RGB(146, 208, 80)
Else
CSR.Visible = xlSheetHidden
Range("J41").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J41").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub DFMEA_box_Click()

If DFMEA_box.Value = True Then
DFMEA.Visible = xlSheetVisible
FMEAR.Visible = xlSheetVisible
Range("J13").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J13").Interior.Color = RGB(146, 208, 80)
Else
DFMEA.Visible = xlSheetHidden
FMEAR.Visible = xlSheetHidden
Range("J13").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J13").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub DR_box_Click()

If DR_box.Value = True Then
DR.Visible = xlSheetVisible
Range("J7").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J7").Interior.Color = RGB(146, 208, 80)
Else
DR.Visible = xlSheetHidden
Range("J7").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J7").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub ECN_box_Click()

If ECN_box.Value = True Then
ECN.Visible = xlSheetVisible
Range("J9").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J9").Interior.Color = RGB(146, 208, 80)
Else
ECN.Visible = xlSheetHidden
Range("J9").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J9").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub FAI_box_Click()

If FAI_box.Value = True Then
FAI.Visible = xlSheetVisible
Range("J19").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J19").Interior.Color = RGB(146, 208, 80)
Else
FAI.Visible = xlSheetHidden
Range("J19").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J19").Interior.Color = RGB(255, 0, 0)
End If


End Sub

Private Sub IPS_box_Click()

If IPS_box.Value = True Then
IPSI.Visible = xlSheetVisible
SPPC.Visible = xlSheetVisible
Range("J23").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J23").Interior.Color = RGB(146, 208, 80)
Else
IPSI.Visible = xlSheetHidden
SPPC.Visible = xlSheetHidden
Range("J23").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J23").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub LCA_box_Click()

If LCA_box.Value = True Then
LCA.Visible = xlSheetVisible
Range("J39").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J39").Interior.Color = RGB(146, 208, 80)
Else
LCA.Visible = xlSheetHidden
Range("J39").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J39").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub MS_box_Click()

If MS_box.Value = True Then
MS.Visible = xlSheetVisible
Range("J37").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J37").Interior.Color = RGB(146, 208, 80)
Else
MS.Visible = xlSheetHidden
Range("J37").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J37").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub MSA_box_Click()

If MSA_box.Value = True Then
MSAA.Visible = xlSheetVisible
MSAG.Visible = xlSheetVisible
Range("J25").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J25").Interior.Color = RGB(146, 208, 80)
Else
MSAA.Visible = xlSheetHidden
MSAG.Visible = xlSheetHidden
Range("J25").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J25").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub MTR_box_Click()

If MTR_box.Value = True Then
MTR.Visible = xlSheetVisible
PTR.Visible = xlSheetVisible
Range("J21").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J21").Interior.Color = RGB(146, 208, 80)
Else
MTR.Visible = xlSheetHidden
PTR.Visible = xlSheetHidden
Range("J21").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J21").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub PFD_box_Click()

If PFD_box.Value = True Then
PFD.Visible = xlSheetVisible
Range("J15").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J15").Interior.Color = RGB(146, 208, 80)
Else
PFD.Visible = xlSheetHidden
Range("J15").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J15").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub PFMEA_box_Click()

If PFMEA_box.Value = True Then
PFMEA.Visible = xlSheetVisible
FMEAR.Visible = xlSheetVisible
Range("J17").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J17").Interior.Color = RGB(146, 208, 80)
Else
PFMEA.Visible = xlSheetHidden
FMEAR.Visible = xlSheetHidden
Range("J17").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J17").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub PSW_box_Click()

If PSW_box.Value = True Then
PSW.Visible = xlSheetVisible
Range("J31").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J31").Interior.Color = RGB(146, 208, 80)
Else
PSW.Visible = xlSheetHidden
Range("J31").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J31").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub QLD_box_Click()

If QLD_box.Value = True Then
QLD.Visible = xlSheetVisible
Range("J27").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J27").Interior.Color = RGB(146, 208, 80)
Else
QLD.Visible = xlSheetHidden
Range("J27").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J27").Interior.Color = RGB(255, 0, 0)
End If

End Sub

Private Sub SP_box_Click()

If SP_box.Value = True Then
SP.Visible = xlSheetVisible
Range("J35").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J35").Interior.Color = RGB(146, 208, 80)
Else
SP.Visible = xlSheetHidden
Range("J35").Interior.Pattern = xlNone 'Remove any manual formatting
Range("J35").Interior.Color = RGB(255, 0, 0)
End If

End Sub



Private Sub optLevel1_Click()

If optLevel1.Value = True Then
DR_box.Value = True
ECN_box.Value = False
CEA_box.Value = False
DFMEA_box.Value = False
PFD_box.Value = False
PFMEA_box.Value = False
FAI_box.Value = False
MTR_box.Value = False
IPS_box.Value = False
MSA_box.Value = False
QLD_box.Value = False
CP_box.Value = False
PSW_box.Value = True
AAR_box.Value = True
SP_box.Value = False
MS_box.Value = False
LCA_box.Value = False
CSR_box.Value = False


End If


End Sub

Private Sub optLevel2_Click()

If optLevel2.Value = True Then
DR_box.Value = True
ECN_box.Value = False
CEA_box.Value = True
DFMEA_box.Value = False
PFD_box.Value = False
PFMEA_box.Value = False
FAI_box.Value = True
MTR_box.Value = True
IPS_box.Value = False
MSA_box.Value = False
QLD_box.Value = True
CP_box.Value = False
PSW_box.Value = True
AAR_box.Value = False
SP_box.Value = False
MS_box.Value = False
LCA_box.Value = False
CSR_box.Value = False
End If

End Sub

Private Sub optLevel3_Click()

If optLevel3.Value = True Then
DR_box.Value = True
ECN_box.Value = True
CEA_box.Value = True
DFMEA_box.Value = True
PFD_box.Value = True
PFMEA_box.Value = True
FAI_box.Value = True
MTR_box.Value = True
IPS_box.Value = True
MSA_box.Value = True
QLD_box.Value = True
CP_box.Value = True
PSW_box.Value = True
AAR_box.Value = True
SP_box.Value = True
MS_box.Value = True
LCA_box.Value = True
CSR_box.Value = True
End If

End Sub

Private Sub optLevel4_Click()

If optLevel4.Value = True Then
DR_box.Value = True
ECN_box.Value = False
CEA_box.Value = False
DFMEA_box.Value = False
PFD_box.Value = False
PFMEA_box.Value = False
FAI_box.Value = False
MTR_box.Value = False
IPS_box.Value = False
MSA_box.Value = False
QLD_box.Value = False
CP_box.Value = False
PSW_box.Value = True
AAR_box.Value = False
SP_box.Value = False
MS_box.Value = False
LCA_box.Value = False
CSR_box.Value = False
End If

End Sub

Private Sub optLevel5_Click()

If optLevel5.Value = True Then
DR_box.Value = False
ECN_box.Value = False
CEA_box.Value = False
DFMEA_box.Value = False
PFD_box.Value = False
PFMEA_box.Value = False
FAI_box.Value = False
MTR_box.Value = False
IPS_box.Value = False
MSA_box.Value = False
QLD_box.Value = False
CP_box.Value = False
PSW_box.Value = False
AAR_box.Value = False
SP_box.Value = False
MS_box.Value = False
LCA_box.Value = False
CSR_box.Value = False
End If

End Sub

