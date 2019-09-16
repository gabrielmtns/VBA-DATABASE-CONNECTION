Attribute VB_Name = "baseFiltrosPivotTable"
Option Explicit


Sub filtrarSegmentacao(ByVal funcionario As String, ByVal status As String)
    Dim meuSlicer As New SlicerControl
    Dim teste As Variant
    
    meuSlicer.pastaDeTrabalho = ActiveWorkbook
    meuSlicer.segmentacaoes "Segmenta��odeDados_baseRepresentantesAtendimento.Funcionario", _
                            "[relatorioCompleto].[baseRepresentantesAtendimento.Funcionario]", _
                            "BIANKA", _
                            "Segmenta��odeDados_BaseDeStatus.STATUS_FINAL", _
                            "[relatorioCompleto].[BaseDeStatus.STATUS FINAL]", _
                            "RESERVADO"
End Sub


